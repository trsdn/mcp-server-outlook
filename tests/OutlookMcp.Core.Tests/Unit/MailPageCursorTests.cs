using OutlookMcp.Core.Commands.Mail;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Tests for the opaque continuation token that <c>mail.list</c> and <c>mail.search</c> hand back
/// so a caller can walk a folder past the first page (#43).
///
/// This is pure encode/decode/validate with zero COM dependency - no Outlook type appears in
/// <see cref="MailPageCursor"/>'s signature - so it falls under the exception Rule 30 carves out,
/// the same one <c>MailRestrictFilterTests</c> and <c>OutlookDispatcherTests</c> document. Walking
/// a real folder across several pages is covered separately by the integration suite.
///
/// Two properties matter more than the encoding details:
///
/// <list type="bullet">
/// <item>A cursor is <b>bound to the query that produced it</b>. Continuing query A with a cursor
/// minted by query B would return a page of the wrong result set while looking perfectly
/// successful, so it is rejected.</item>
/// <item>An unusable cursor <b>fails loudly</b>. Silently restarting from the first page would
/// re-emit page one forever and never surface the rest of the folder - a caller looping until
/// <c>nextCursor</c> is null would hang, or worse, stop early believing it had seen everything.</item>
/// </list>
/// </summary>
public class MailPageCursorTests
{
    private const string Fingerprint = "inbox|q=budget";

    private static readonly DateTimeOffset Boundary =
        new(2024, 3, 7, 14, 30, 0, TimeSpan.Zero);

    [Fact]
    public void Encode_ThenDecode_RoundTripsTheBoundary()
    {
        string token = MailPageCursor.Encode(Fingerprint, Boundary, ["entry-1"]);

        Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out string? error));
        Assert.Null(error);
        Assert.NotNull(cursor);
        Assert.Equal(Boundary, cursor!.LastReceived);
    }

    [Fact]
    public void Encode_ThenDecode_RoundTripsTheBoundaryEntryIds()
    {
        // The ids seen at exactly the boundary timestamp are what makes paging exact when several
        // messages share a received time: the next page re-scans that band and skips these by id
        // rather than trusting Outlook to order ties the same way twice.
        string token = MailPageCursor.Encode(Fingerprint, Boundary, ["entry-1", "entry-2"]);

        Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out _));

        Assert.Equal(["entry-1", "entry-2"], cursor!.SeenAtBoundary);
    }

    [Fact]
    public void Encode_PreservesSubSecondPrecision()
    {
        // Truncating to whole seconds would widen the boundary band, and every message inside that
        // band that was not recorded by id would be silently skipped on the next page.
        var precise = new DateTimeOffset(2024, 3, 7, 14, 30, 0, 123, TimeSpan.Zero);

        string token = MailPageCursor.Encode(Fingerprint, precise, ["entry-1"]);

        Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out _));
        Assert.Equal(precise, cursor!.LastReceived);
    }

    [Fact]
    public void Encode_NormalisesTheBoundaryToUtc()
    {
        // Outlook compares urn:schemas:httpmail:datereceived in UTC. A cursor carrying local
        // wall-clock time would shift the next page's lower bound by the offset and drop a band of
        // mail - the same bug the Restrict date filters had (#27).
        var local = new DateTimeOffset(2024, 3, 7, 16, 30, 0, TimeSpan.FromHours(2));

        string token = MailPageCursor.Encode(Fingerprint, local, []);

        Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out _));
        Assert.Equal(TimeSpan.Zero, cursor!.LastReceived.Offset);
        Assert.Equal(local.UtcDateTime, cursor.LastReceived.UtcDateTime);
    }

    [Fact]
    public void Encode_ProducesAUrlSafeToken()
    {
        // The token travels through CLI arguments and JSON. Base64 '+' and '/' survive neither
        // comfortably, and '=' padding invites shell mangling.
        string token = MailPageCursor.Encode(Fingerprint, Boundary, ["entry/with+odd=chars"]);

        Assert.DoesNotContain('+', token);
        Assert.DoesNotContain('/', token);
        Assert.DoesNotContain('=', token);
    }

    [Fact]
    public void TryDecode_WithACursorFromADifferentQuery_Fails()
    {
        string token = MailPageCursor.Encode(Fingerprint, Boundary, ["entry-1"]);

        bool ok = MailPageCursor.TryDecode(token, "inbox|q=invoices", out MailPageCursor? cursor, out string? error);

        Assert.False(ok);
        Assert.Null(cursor);
        Assert.NotNull(error);
    }

    [Fact]
    public void TryDecode_WithACursorFromADifferentQuery_SaysSo()
    {
        string token = MailPageCursor.Encode(Fingerprint, Boundary, ["entry-1"]);

        MailPageCursor.TryDecode(token, "inbox|q=invoices", out _, out string? error);

        // The caller has to be able to tell "your cursor is stale, start over" apart from
        // "your cursor is corrupt", because only one of those is their fault.
        Assert.Contains("different", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not-a-cursor")]
    [InlineData("!!!!")]
    public void TryDecode_WithAMalformedToken_FailsInsteadOfRestarting(string token)
    {
        bool ok = MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out string? error);

        Assert.False(ok);
        Assert.Null(cursor);
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void TryDecode_WithATamperedPayload_Fails()
    {
        string token = MailPageCursor.Encode(Fingerprint, Boundary, ["entry-1"]);
        string tampered = token[..^4] + "AAAA";

        bool ok = MailPageCursor.TryDecode(tampered, Fingerprint, out _, out string? error);

        Assert.False(ok);
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void TryDecode_WithAFutureVersion_Fails()
    {
        // A token minted by a newer build may mean something different by the same fields. Refusing
        // it is the only safe reading; guessing risks a page that quietly skips mail.
        string forged = MailPageCursor.EncodeForTest(version: 99, Fingerprint, Boundary, ["entry-1"]);

        bool ok = MailPageCursor.TryDecode(forged, Fingerprint, out _, out string? error);

        Assert.False(ok);
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void Fingerprint_DiffersWhenTheFolderDiffers()
    {
        string a = MailPageCursor.BuildFingerprint("inbox", null, false, null, null, null, null, null);
        string b = MailPageCursor.BuildFingerprint("drafts", null, false, null, null, null, null, null);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void Fingerprint_DiffersWhenTheQueryDiffers()
    {
        string a = MailPageCursor.BuildFingerprint("inbox", "budget", false, null, null, null, null, null);
        string b = MailPageCursor.BuildFingerprint("inbox", "invoices", false, null, null, null, null, null);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void Fingerprint_DiffersWhenAStructuredFilterDiffers()
    {
        // maxCount is deliberately excluded from the fingerprint - changing page size mid-walk is
        // legitimate - but a filter change alters which items exist, so it must invalidate.
        string a = MailPageCursor.BuildFingerprint("inbox", null, false, "a@example.com", null, null, null, null);
        string b = MailPageCursor.BuildFingerprint("inbox", null, false, "b@example.com", null, null, null, null);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void Fingerprint_IsStableAcrossCalls()
    {
        string a = MailPageCursor.BuildFingerprint("inbox", "budget", true, "a@example.com", "re:", "2024-01-01", null, true);
        string b = MailPageCursor.BuildFingerprint("inbox", "budget", true, "a@example.com", "re:", "2024-01-01", null, true);

        Assert.Equal(a, b);
    }

    [Fact]
    public void Fingerprint_TreatsNullFolderAsItsOwnCase()
    {
        // Null folder means "the active/default folder", which is not the same query as an
        // explicitly named one even if they happen to resolve to the same place today.
        string a = MailPageCursor.BuildFingerprint(null, null, false, null, null, null, null, null);
        string b = MailPageCursor.BuildFingerprint("inbox", null, false, null, null, null, null, null);

        Assert.NotEqual(a, b);
    }

    // ---------------------------------------------------------------------------------------
    // Boundary decisions.
    //
    // These live here rather than in the integration suite for a reason worth recording. The
    // paging integration tests page over a real folder and passed unchanged when the boundary
    // comparison was mutated from '>' to '>=' - a change that silently drops tied items - because
    // no two messages in that folder share a received time to the millisecond. The integration
    // tests are still the ones that prove the COM walk works, but the tie logic is decision logic
    // with zero COM dependency and has to be pinned down exhaustively here, where the awkward
    // cases can actually be constructed.
    // ---------------------------------------------------------------------------------------

    private static MailPageCursor CursorAt(DateTimeOffset boundary, params string[] seen)
    {
        string token = MailPageCursor.Encode(Fingerprint, boundary, seen);
        Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out _));
        return cursor!;
    }

    [Fact]
    public void Includes_ItemNewerThanTheBoundary_IsExcluded()
    {
        // It was served on an earlier page.
        var cursor = CursorAt(Boundary, "entry-1");

        Assert.False(cursor.Includes(Boundary.AddSeconds(1), "entry-newer"));
    }

    [Fact]
    public void Includes_ItemOlderThanTheBoundary_IsIncluded()
    {
        var cursor = CursorAt(Boundary, "entry-1");

        Assert.True(cursor.Includes(Boundary.AddSeconds(-1), "entry-older"));
    }

    [Fact]
    public void Includes_ItemAtTheBoundaryAlreadyServed_IsExcluded()
    {
        var cursor = CursorAt(Boundary, "entry-1");

        Assert.False(cursor.Includes(Boundary, "entry-1"));
    }

    [Fact]
    public void Includes_ItemAtTheBoundaryNotYetServed_IsIncluded()
    {
        // The case a '>=' comparison silently drops. Two messages delivered in the same batch can
        // share a received time to the millisecond; excluding the whole instant because one of them
        // was already returned loses the other permanently, and the caller is never told.
        var cursor = CursorAt(Boundary, "entry-1");

        Assert.True(cursor.Includes(Boundary, "entry-2"));
    }

    [Fact]
    public void Includes_ItemAtTheBoundaryWithNoEntryId_IsIncluded()
    {
        // An unreadable id must not cause an item to vanish. Serving it twice is visible to the
        // caller; dropping it is not.
        var cursor = CursorAt(Boundary, "entry-1");

        Assert.True(cursor.Includes(Boundary, null));
    }

    [Fact]
    public void Includes_ComparesTheBoundaryInUtcNotWallClock()
    {
        // Same instant, expressed in a different offset. Treating these as different would shift
        // the frontier by the offset and drop a band of mail - the bug the Restrict date filters
        // had (#27).
        var cursor = CursorAt(Boundary, "entry-1");
        var sameInstantElsewhere = Boundary.ToOffset(TimeSpan.FromHours(2));

        Assert.False(cursor.Includes(sameInstantElsewhere, "entry-1"));
        Assert.True(cursor.Includes(sameInstantElsewhere, "entry-2"));
    }

    [Fact]
    public void Includes_WithSeveralTiedItemsServed_ExcludesOnlyThoseServed()
    {
        var cursor = CursorAt(Boundary, "entry-1", "entry-2", "entry-3");

        Assert.False(cursor.Includes(Boundary, "entry-2"));
        Assert.True(cursor.Includes(Boundary, "entry-4"));
    }

    [Fact]
    public void Includes_MatchesEntryIdsCaseSensitively()
    {
        // Outlook entry ids are hex strings compared by identity. Case-insensitive matching would
        // risk excluding an item that merely looks similar, which is a silent omission.
        var cursor = CursorAt(Boundary, "ABCD");

        Assert.True(cursor.Includes(Boundary, "abcd"));
    }
}
