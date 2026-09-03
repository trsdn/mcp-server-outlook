using System.Buffers.Text;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// The opaque continuation token that lets a caller walk <c>mail.list</c> / <c>mail.search</c>
/// results past the first page (#43).
///
/// <para>
/// Before this existed, a response could only say <c>truncated: true</c>. That tells a caller there
/// is more mail but gives them no way to reach it: re-issuing the same call returns the same first
/// page forever, and there was no offset to advance. The honest reading of a truncated response was
/// "some matches exist and you cannot see the rest", which for an agent-facing tool is barely better
/// than a wrong answer.
/// </para>
///
/// <para>
/// <b>Why a keyset cursor rather than an offset.</b> Results are ordered by received time,
/// descending. An offset ("resume at item 26") is wrong the moment mail arrives: every item shifts
/// down by one, so item 26 on the second call is item 25 from the first, and the caller sees a
/// duplicate while a different message is skipped entirely. This cursor instead records the
/// <i>received time of the last item returned</i>, so the next page continues from a point in the
/// ordering rather than a position in a list. New mail arriving above the boundary does not disturb
/// it.
/// </para>
///
/// <para>
/// <b>Why the boundary carries entry ids.</b> Received times are not unique - a batch delivered
/// together can share a timestamp to the millisecond. Continuing from "strictly older than the
/// boundary" would skip the rest of that band; continuing from "older than or equal" would repeat
/// the ones already sent. So the cursor also carries the ids returned <i>at exactly</i> the boundary
/// instant, the next page re-scans that band inclusively, and those ids are skipped by identity.
/// Ties are rare, so this stays small, and it does not assume Outlook orders ties the same way
/// twice.
/// </para>
///
/// <para>
/// <b>Why it is bound to the query.</b> A cursor minted by one query would otherwise continue a
/// different one, returning a page of the wrong result set with every appearance of success. The
/// fingerprint makes that a clean error instead.
/// </para>
///
/// <para>
/// <b>What this does not promise.</b> It is a keyset cursor over a live mailbox, not a snapshot.
/// Mail deleted mid-walk will not appear, and mail that arrives mid-walk with a received time above
/// the boundary is not retro-fitted into a page already passed. What it does guarantee is the part
/// that matters: within the range walked, no item is returned twice and none is silently skipped.
/// </para>
/// </summary>
internal sealed class MailPageCursor
{
    private const int CurrentVersion = 1;

    /// <summary>Received time of the last item on the previous page, always UTC.</summary>
    public DateTimeOffset LastReceived { get; private init; }

    /// <summary>
    /// Entry ids already returned at exactly <see cref="LastReceived"/>. The next page re-scans that
    /// instant and skips these, which is what makes tied received times safe.
    /// </summary>
    public IReadOnlyList<string> SeenAtBoundary { get; private init; } = [];

    /// <summary>
    /// Decides whether an item sits past this cursor's frontier, i.e. whether it belongs to a page
    /// that has not been served yet.
    ///
    /// <para>
    /// Items newer than the boundary were returned on an earlier page. Items at exactly the
    /// boundary instant are ambiguous - received times are not unique - so that band is re-scanned
    /// and the ids already served are excluded by identity, rather than by trusting Outlook to
    /// order tied items the same way on a second call.
    /// </para>
    ///
    /// <para>
    /// The comparison against the boundary is deliberately <c>&gt;</c> and not <c>&gt;=</c>. Excluding
    /// the whole boundary instant would drop every tied item that had <i>not</i> already been
    /// served - an unrecoverable silent omission, and the exact failure this paging contract exists
    /// to prevent.
    /// </para>
    ///
    /// <para>
    /// An item whose received time could not be read is <b>kept</b>, for the same reason: at worst
    /// it is served twice, which a caller can see and handle, whereas dropping it is invisible.
    /// </para>
    /// </summary>
    public bool Includes(DateTimeOffset receivedUtc, string? entryId)
    {
        if (receivedUtc > LastReceived)
        {
            return false;
        }

        return receivedUtc != LastReceived
               || entryId is null
               || !SeenAtBoundary.Contains(entryId, StringComparer.Ordinal);
    }

    /// <summary>
    /// Builds the query fingerprint a cursor is bound to. Deliberately excludes <c>maxCount</c>:
    /// changing page size part-way through a walk is legitimate and does not change which items
    /// exist. Everything that <i>does</i> change the result set is included.
    /// </summary>
    public static string BuildFingerprint(
        string? folder,
        string? query,
        bool unreadOnly,
        string? fromAddress,
        string? subjectContains,
        string? receivedAfter,
        string? receivedBefore,
        bool? hasAttachment,
        string? searchMode = null)
    {
        // Unit separator: cannot occur in any of these values, so no field can impersonate another
        // by containing the delimiter.
        string raw = string.Join(
            '\u001f',
            folder ?? "\u0000",
            query ?? "\u0000",
            unreadOnly ? "1" : "0",
            fromAddress ?? "\u0000",
            subjectContains ?? "\u0000",
            receivedAfter ?? "\u0000",
            receivedBefore ?? "\u0000",
            hasAttachment?.ToString(CultureInfo.InvariantCulture) ?? "\u0000",
            // The engine is part of the query's identity: the same words answered by the content
            // index and by the client-side scan are different result sets, so a cursor minted by one
            // must not be accepted by the other.
            searchMode ?? "\u0000");

        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(raw)))[..16];
    }

    /// <summary>
    /// Mints a token for the next page. <paramref name="seenAtBoundary"/> must be every entry id
    /// returned at exactly <paramref name="lastReceived"/> on this page and any earlier page that
    /// shared the instant - omitting one causes it to be re-emitted.
    /// </summary>
    public static string Encode(
        string fingerprint,
        DateTimeOffset lastReceived,
        IReadOnlyList<string> seenAtBoundary)
        => EncodeCore(CurrentVersion, fingerprint, lastReceived, seenAtBoundary);

    /// <summary>
    /// Mints a token with an explicit version so tests can prove an unrecognised version is
    /// rejected rather than best-guessed. Not used by production code.
    /// </summary>
    internal static string EncodeForTest(
        int version,
        string fingerprint,
        DateTimeOffset lastReceived,
        IReadOnlyList<string> seenAtBoundary)
        => EncodeCore(version, fingerprint, lastReceived, seenAtBoundary);

    private static string EncodeCore(
        int version,
        string fingerprint,
        DateTimeOffset lastReceived,
        IReadOnlyList<string> seenAtBoundary)
    {
        var payload = new CursorPayload
        {
            Version = version,
            Fingerprint = fingerprint,
            // "O" on a UTC-normalised value keeps sub-second precision. Truncating to whole seconds
            // would widen the boundary band beyond the ids recorded for it, and everything else in
            // that widened band would be skipped.
            LastReceived = lastReceived.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture),
            SeenAtBoundary = [.. seenAtBoundary]
        };

        byte[] json = JsonSerializer.SerializeToUtf8Bytes(payload, CursorJson);
        return Base64Url.EncodeToString(json);
    }

    /// <summary>
    /// Decodes and validates a token. Returns <see langword="false"/> with a caller-facing
    /// <paramref name="error"/> rather than throwing, and <b>never</b> falls back to "start from the
    /// beginning" - a silent restart would re-emit page one indefinitely, so a caller looping until
    /// <c>nextCursor</c> is null would never terminate and one checking for new results would
    /// conclude the folder held nothing further.
    /// </summary>
    public static bool TryDecode(
        string? token,
        string expectedFingerprint,
        out MailPageCursor? cursor,
        out string? error)
    {
        cursor = null;
        error = null;

        if (string.IsNullOrWhiteSpace(token))
        {
            error = "cursor is empty. Omit it to start from the first page.";
            return false;
        }

        CursorPayload? payload;
        try
        {
            byte[] json = Base64Url.DecodeFromChars(token.AsSpan());
            payload = JsonSerializer.Deserialize<CursorPayload>(json, CursorJson);
        }
        catch (Exception ex) when (ex is FormatException or JsonException or DecoderFallbackException)
        {
            error = "cursor is malformed. Pass back the nextCursor value exactly as it was returned, "
                  + "or omit it to start from the first page.";
            return false;
        }

        if (payload is null)
        {
            error = "cursor is malformed. Pass back the nextCursor value exactly as it was returned, "
                  + "or omit it to start from the first page.";
            return false;
        }

        if (payload.Version != CurrentVersion)
        {
            error = $"cursor version {payload.Version} is not supported by this build "
                  + $"(expected {CurrentVersion}). Restart the listing without a cursor.";
            return false;
        }

        if (!string.Equals(payload.Fingerprint, expectedFingerprint, StringComparison.Ordinal))
        {
            error = "cursor belongs to a different query. A cursor can only continue the exact "
                  + "folder, query and filters that produced it; changing any of them requires "
                  + "restarting without a cursor.";
            return false;
        }

        if (!DateTimeOffset.TryParse(
                payload.LastReceived,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out DateTimeOffset lastReceived))
        {
            error = "cursor carries an unreadable position. Restart the listing without a cursor.";
            return false;
        }

        cursor = new MailPageCursor
        {
            LastReceived = lastReceived.ToUniversalTime(),
            SeenAtBoundary = payload.SeenAtBoundary ?? []
        };
        return true;
    }

    private static readonly JsonSerializerOptions CursorJson = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private sealed class CursorPayload
    {
        [JsonPropertyName("v")]
        public int Version { get; set; }

        [JsonPropertyName("q")]
        public string? Fingerprint { get; set; }

        [JsonPropertyName("t")]
        public string? LastReceived { get; set; }

        [JsonPropertyName("e")]
        public List<string>? SeenAtBoundary { get; set; }
    }
}
