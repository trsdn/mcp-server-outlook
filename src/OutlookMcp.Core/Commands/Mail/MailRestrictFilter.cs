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
/// </summary>
internal static class MailRestrictFilter
{
    private const string ReadProperty = "urn:schemas:httpmail:read";
    private const string SubjectProperty = "urn:schemas:httpmail:subject";
    private const string FromEmailProperty = "urn:schemas:httpmail:fromemail";
    private const string FromNameProperty = "urn:schemas:httpmail:fromname";
    private const string DateReceivedProperty = "urn:schemas:httpmail:datereceived";
    private const string HasAttachmentProperty = "urn:schemas:httpmail:hasattachment";

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
        bool? hasAttachment = null)
    {
        var clauses = new List<string>();

        if (unreadOnly)
        {
            clauses.Add($"{Quote(ReadProperty)} = 0");
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

        return clauses.Count == 0
            ? null
            : "@SQL=" + string.Join(" AND ", clauses);
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
