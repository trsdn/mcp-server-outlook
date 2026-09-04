using System.Globalization;

namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// Builds the DASL filter string passed to <c>Items.Restrict</c> so that structured mail predicates
/// are evaluated by Outlook/MAPI rather than by hydrating and inspecting every item client-side
/// (#27).
///
/// <para>
/// The controlling design rule is that the filter must never be <b>under</b>-inclusive. Outlook
/// applies it before the client sees anything, so an item this filter wrongly excludes cannot be
/// recovered by the client-side narrowing that follows: the caller is simply told the mail does not
/// exist. A silent false negative is the worst outcome for an agent-facing tool, so wherever a
/// predicate cannot be expressed exactly in DASL it is dropped entirely and left to the client-side
/// check. That costs speed, never correctness.
/// </para>
///
/// <para>
/// Escaping follows the documented DASL rules: a value is delimited by single quotes and an
/// embedded single quote is escaped by doubling it. Getting this wrong does not merely fail to
/// match - it yields a syntactically invalid filter that <c>Restrict</c> rejects outright.
/// See <see href="https://learn.microsoft.com/office/vba/outlook/how-to/search-and-filter/filtering-items-using-a-string-comparison"/>.
/// </para>
///
/// <para>
/// <b>The full-text clause is the one exception to that rule, and it is why it is opt-in (#42).</b>
/// <c>ci_phrasematch</c> asks the content index, which matches whole words: it finds <c>foo</c> in
/// "a foo arrived" but not inside "foobar", where the client-side substring check would. So it is
/// not a drop-in speed-up for the free-text scan - it answers a slightly different question, faster
/// and without the scan's horizon. A caller must ask for it, and the response names which engine
/// answered, because an empty result means different things depending on which one did.
/// </para>
/// </summary>
internal static class MailRestrictFilter
{
    private const string ReadProperty = "urn:schemas:httpmail:read";
    private const string SubjectProperty = "urn:schemas:httpmail:subject";
    private const string FromEmailProperty = "urn:schemas:httpmail:fromemail";
    private const string FromNameProperty = "urn:schemas:httpmail:fromname";
    private const string DateReceivedProperty = "urn:schemas:httpmail:datereceived";
    private const string HasAttachmentProperty = "urn:schemas:httpmail:hasattachment";
    private const string BodyProperty = "urn:schemas:httpmail:textdescription";
    private const string DisplayToProperty = "urn:schemas:httpmail:displayto";
    private const string DisplayCcProperty = "urn:schemas:httpmail:displaycc";

    /// <summary>
    /// <c>PR_FLAG_STATUS</c>. There is no <c>urn:schemas:httpmail:</c> equivalent that carries the
    /// follow-up state, so this is addressed by MAPI property tag: <c>0x1090</c> with type
    /// <c>PT_LONG</c> (<c>0003</c>). Verified against a live mailbox rather than inferred - a
    /// mis-named property makes <c>Restrict</c> match nothing and reports it as "no such mail".
    /// </summary>
    private const string FlagStatusProperty = "http://schemas.microsoft.com/mapi/proptag/0x10900003";

    /// <summary><c>olFlagMarked</c> - a follow-up that is still outstanding.</summary>
    private const int FlagMarked = 2;

    /// <summary>
    /// The fields a full-text query is asked against, in the order they are emitted.
    ///
    /// <para>
    /// This list mirrors the client-side free-text check exactly. Pushing only the body down would be
    /// under-inclusive in the worst way: Outlook would discard a subject or sender match before the
    /// client ever saw the item, and the caller would be told the mail does not exist.
    /// </para>
    /// </summary>
    private static readonly string[] FullTextProperties =
    [
        BodyProperty,
        SubjectProperty,
        FromNameProperty,
        FromEmailProperty,
        DisplayToProperty,
        DisplayCcProperty
    ];

    /// <summary>
    /// Builds an <c>@SQL=</c> DASL filter for the supplied predicates, or <see langword="null"/>
    /// when none of them are set - in which case the caller should skip <c>Restrict</c> altogether
    /// rather than issue a match-everything filter.
    /// </summary>
    public static string? Build(
        bool unreadOnly = false,
        string? fromAddress = null,
        string? subjectContains = null,
        DateTimeOffset? receivedAfter = null,
        DateTimeOffset? receivedBefore = null,
        bool? hasAttachment = null,
        bool flaggedOnly = false,
        string? fullTextQuery = null)
    {
        List<string> clauses = BuildStructuredClauses(
            unreadOnly, fromAddress, subjectContains, receivedAfter, receivedBefore, hasAttachment, flaggedOnly);

        // Content-index full-text (#42). Deliberately last so the cheap structured predicates are
        // written first; Outlook is free to reorder, but the emitted string reads the way a human
        // would debug it.
        if (!string.IsNullOrWhiteSpace(fullTextQuery))
        {
            string escaped = EscapeLiteral(fullTextQuery.Trim());
            clauses.Add(
                "(" + string.Join(
                    " OR ",
                    FullTextProperties.Select(property => $"{Quote(property)} ci_phrasematch '{escaped}'"))
                + ")");
        }

        return clauses.Count == 0
            ? null
            : "@SQL=" + string.Join(" AND ", clauses);
    }

    /// <summary>
    /// Builds the filter for <c>Application.AdvancedSearch</c> (#13).
    ///
    /// <para>
    /// Two things differ from <see cref="Build"/>, and both are load bearing.
    /// </para>
    ///
    /// <para>
    /// <b>No <c>@SQL=</c> prefix.</b> <c>Items.Restrict</c> requires it to select the DASL dialect;
    /// <c>AdvancedSearch</c> is DASL-only and rejects the prefixed form outright with a bare
    /// "The operation failed." - verified against a live mailbox, not inferred.
    /// </para>
    ///
    /// <para>
    /// <b><c>LIKE</c> rather than <c>ci_phrasematch</c> for the free-text clause.</b> This is the
    /// engine that replaces the client-side scan, so it has to answer the same question the scan
    /// did: substring matching. <c>ci_phrasematch</c> asks the content index, which matches whole
    /// words - swapping one for the other would silently stop finding every mid-word match while the
    /// response still said the search succeeded. Callers who want the index ask for it by name.
    /// </para>
    ///
    /// <para>
    /// Returns <see langword="null"/> when the free-text term cannot be expressed exactly - a
    /// wildcard the caller supplied, since DASL has no <c>ESCAPE</c> clause. The caller falls back to
    /// an engine that can answer it rather than running a quietly widened search.
    /// </para>
    /// </summary>
    public static string? BuildAdvancedSearch(
        string freeTextQuery,
        bool unreadOnly = false,
        string? fromAddress = null,
        string? subjectContains = null,
        DateTimeOffset? receivedAfter = null,
        DateTimeOffset? receivedBefore = null,
        bool? hasAttachment = null,
        bool flaggedOnly = false)
    {
        if (!IsUsableLikeValue(freeTextQuery))
        {
            return null;
        }

        List<string> clauses = BuildStructuredClauses(
            unreadOnly, fromAddress, subjectContains, receivedAfter, receivedBefore, hasAttachment, flaggedOnly);

        string escaped = EscapeLiteral(freeTextQuery.Trim());
        clauses.Add(
            "(" + string.Join(
                " OR ",
                FullTextProperties.Select(property => $"{Quote(property)} LIKE '%{escaped}%'"))
            + ")");

        return string.Join(" AND ", clauses);
    }

    private static List<string> BuildStructuredClauses(
        bool unreadOnly,
        string? fromAddress,
        string? subjectContains,
        DateTimeOffset? receivedAfter,
        DateTimeOffset? receivedBefore,
        bool? hasAttachment,
        bool flaggedOnly)
    {
        var clauses = new List<string>();

        if (unreadOnly)
        {
            clauses.Add($"{Quote(ReadProperty)} = 0");
        }

        // Deliberately "= 2" and not "<> 0". A completed flag is finished work, and returning it
        // under "flagged" would put items the user has already dealt with back on their list.
        if (flaggedOnly)
        {
            clauses.Add($"{Quote(FlagStatusProperty)} = {FlagMarked}");
        }

        if (TryBuildContains(SubjectProperty, subjectContains, out string? subjectClause))
        {
            clauses.Add(subjectClause);
        }

        // A caller filtering "from: alice" does not necessarily know whether that is an address or
        // a display name, so both are matched. Choosing one would manufacture false negatives.
        if (IsUsableLikeValue(fromAddress))
        {
            string escaped = EscapeLiteral(fromAddress!.Trim());
            clauses.Add(
                $"({Quote(FromEmailProperty)} LIKE '%{escaped}%' " +
                $"OR {Quote(FromNameProperty)} LIKE '%{escaped}%')");
        }

        if (receivedAfter.HasValue)
        {
            clauses.Add($"{Quote(DateReceivedProperty)} >= '{FormatDate(receivedAfter.Value, -1)}'");
        }

        if (receivedBefore.HasValue)
        {
            clauses.Add($"{Quote(DateReceivedProperty)} <= '{FormatDate(receivedBefore.Value, 1)}'");
        }

        if (hasAttachment.HasValue)
        {
            clauses.Add($"{Quote(HasAttachmentProperty)} = {(hasAttachment.Value ? 1 : 0)}");
        }

        return clauses;
    }

    private static bool TryBuildContains(string property, string? value, out string clause)
    {
        if (!IsUsableLikeValue(value))
        {
            clause = string.Empty;
            return false;
        }

        clause = $"{Quote(property)} LIKE '%{EscapeLiteral(value!.Trim())}%'";
        return true;
    }

    /// <summary>
    /// A value is usable in a <c>LIKE</c> pattern only if it contains no wildcard character. DASL
    /// offers no <c>ESCAPE</c> clause, so a literal <c>%</c> or <c>_</c> supplied by the caller
    /// cannot be neutralised and would silently widen or distort the query. Such a predicate is
    /// dropped instead, leaving the search over-inclusive and the client-side check to narrow it.
    /// </summary>
    private static bool IsUsableLikeValue(string? value)
        => !string.IsNullOrWhiteSpace(value)
           && !value.Contains('%', StringComparison.Ordinal)
           && !value.Contains('_', StringComparison.Ordinal);

    private static string EscapeLiteral(string value)
        => value.Replace("'", "''", StringComparison.Ordinal);

    /// <summary>
    /// Formats a date for DASL using the invariant culture, in UTC, with one minute of slack applied
    /// in the widening direction.
    /// <para>
    /// Two things here are load-bearing. First, Outlook compares
    /// <c>urn:schemas:httpmail:datereceived</c> in UTC; a local wall-clock literal makes Restrict
    /// silently drop every message inside the caller's UTC-offset window, and Restrict runs before
    /// the client-side check can recover it. Second, the literal has minute resolution, so a value
    /// carrying seconds would otherwise be truncated toward exclusion on one of the two bounds.
    /// </para>
    /// <para>
    /// The invariant culture matters for a separate reason: a locale-formatted date would make the
    /// same query return different mail on a machine with different regional settings.
    /// </para>
    /// </summary>
    /// <param name="value">The caller-supplied bound.</param>
    /// <param name="slackMinutes">
    /// Negative for a lower bound, positive for an upper bound, so the emitted window is never
    /// narrower than the requested one.
    /// </param>
    private static string FormatDate(DateTimeOffset value, int slackMinutes)
        => value.ToUniversalTime()
                .AddMinutes(slackMinutes)
                .ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

    private static string Quote(string property) => "\"" + property + "\"";
}
