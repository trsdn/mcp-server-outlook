using OutlookMcp.Core.Commands.Mail;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Tests for the DASL filter string built for <c>Items.Restrict</c> (#27).
///
/// This is pure string construction with zero COM dependency - no Outlook type appears in
/// <see cref="MailRestrictFilter"/>'s signature - so it falls under the exception Rule 30 carves
/// out, the same one <c>OutlookDispatcherTests</c> documents. The COM execution path that consumes
/// this string is covered separately by the integration suite.
///
/// The contract under test is deliberately asymmetric: the filter must never be *under*-inclusive.
/// Outlook applies it server-side, so anything it wrongly excludes is invisible to the client-side
/// narrowing that runs afterwards, and the caller sees a false "no such mail exists" - the exact
/// failure #27 exists to remove.
/// </summary>
public class MailRestrictFilterTests
{
    [Fact]
    public void Build_WithNoCriteria_ReturnsNull()
    {
        // No predicates means no Restrict call at all, rather than a filter that matches everything.
        string? filter = MailRestrictFilter.Build();

        Assert.Null(filter);
    }

    /// <summary>
    /// "Flagged" must mean the flag is still outstanding, not merely that a flag was ever raised.
    /// Matching completed items too would answer "what still needs follow-up?" with work the user
    /// has already finished - the one answer that makes the filter worse than not having it.
    /// </summary>
    [Fact]
    public void Build_FlaggedOnly_MatchesOutstandingFlagsAndNotCompletedOnes()
    {
        string? filter = MailRestrictFilter.Build(flaggedOnly: true);

        Assert.Equal(
            "@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" = 2",
            filter);
    }

    [Fact]
    public void Build_FlaggedOnlyFalse_AddsNoClause()
    {
        // Not asking for flagged mail is not the same as asking for unflagged mail.
        string? filter = MailRestrictFilter.Build(flaggedOnly: false);

        Assert.Null(filter);
    }

    [Fact]
    public void Build_FlaggedOnlyWithOtherPredicates_CombinesWithAnd()
    {
        string? filter = MailRestrictFilter.Build(unreadOnly: true, flaggedOnly: true);

        Assert.NotNull(filter);
        Assert.Contains(" AND ", filter, StringComparison.Ordinal);
        Assert.Contains("0x10900003\" = 2", filter, StringComparison.Ordinal);
        Assert.Contains("\"urn:schemas:httpmail:read\" = 0", filter, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_UnreadOnly_UsesDaslReadFlag()
    {
        string? filter = MailRestrictFilter.Build(unreadOnly: true);

        Assert.Equal("@SQL=\"urn:schemas:httpmail:read\" = 0", filter);
    }

    [Fact]
    public void Build_SubjectContains_UsesLikeWithWildcards()
    {
        string? filter = MailRestrictFilter.Build(subjectContains: "invoice");

        Assert.Equal("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%invoice%'", filter);
    }

    [Fact]
    public void Build_ValueContainingApostrophe_DoublesIt()
    {
        // DASL escapes an embedded single quote by doubling it. Getting this wrong does not just
        // fail to match - it produces a syntactically broken filter that Restrict rejects outright.
        string? filter = MailRestrictFilter.Build(subjectContains: "O'Reilly");

        Assert.Equal("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%O''Reilly%'", filter);
    }

    [Theory]
    [InlineData("50% off")]
    [InlineData("draft_v2")]
    [InlineData("100%")]
    public void Build_ValueContainingLikeWildcard_DropsThatPredicate(string value)
    {
        // DASL LIKE has no ESCAPE clause, so % and _ in a user-supplied value cannot be neutralised.
        // Leaving them in would silently change the meaning of the filter. Dropping the predicate
        // keeps the query over-inclusive, and the client-side check still narrows the result
        // exactly - slower, but never wrong.
        string? filter = MailRestrictFilter.Build(subjectContains: value);

        Assert.Null(filter);
    }

    [Fact]
    public void Build_WildcardValueDoesNotDiscardOtherPredicates()
    {
        // Only the unrepresentable predicate is dropped; everything else must still be pushed down.
        string? filter = MailRestrictFilter.Build(unreadOnly: true, subjectContains: "50% off");

        Assert.Equal("@SQL=\"urn:schemas:httpmail:read\" = 0", filter);
    }

    [Fact]
    public void Build_FromAddress_MatchesSenderAddressOrDisplayName()
    {
        // A caller says "from: alice" without knowing whether that is an address or a display name,
        // so both are matched and OR-ed. Narrowing to one of them would produce false negatives.
        string? filter = MailRestrictFilter.Build(fromAddress: "alice@example.com");

        Assert.Equal(
            "@SQL=(\"urn:schemas:httpmail:fromemail\" LIKE '%alice@example.com%' " +
            "OR \"urn:schemas:httpmail:fromname\" LIKE '%alice@example.com%')",
            filter);
    }

    [Fact]
    public void Build_HasAttachmentTrue_UsesDaslFlag()
    {
        string? filter = MailRestrictFilter.Build(hasAttachment: true);

        Assert.Equal("@SQL=\"urn:schemas:httpmail:hasattachment\" = 1", filter);
    }

    [Fact]
    public void Build_HasAttachmentFalse_IsDistinctFromUnset()
    {
        // bool? is tri-state on purpose: false means "only mail without attachments", not "unset".
        string? filter = MailRestrictFilter.Build(hasAttachment: false);

        Assert.Equal("@SQL=\"urn:schemas:httpmail:hasattachment\" = 0", filter);
    }

    [Fact]
    public void Build_ReceivedAfter_FormatsDateInvariantlyAndInclusively()
    {
        var after = new DateTimeOffset(2024, 3, 7, 14, 30, 0, TimeSpan.Zero);

        string? filter = MailRestrictFilter.Build(receivedAfter: after);

        // Invariant format, never the machine's locale: a filter built on a de-DE machine must be
        // byte-identical to one built on en-US, or the same query silently returns different mail.
        // One minute of slack on the lower bound keeps the filter over-inclusive.
        Assert.Equal("@SQL=\"urn:schemas:httpmail:datereceived\" >= '2024-03-07 14:29'", filter);
    }

    [Fact]
    public void Build_ReceivedBefore_IsInclusive()
    {
        var before = new DateTimeOffset(2024, 12, 31, 23, 59, 0, TimeSpan.Zero);

        string? filter = MailRestrictFilter.Build(receivedBefore: before);

        Assert.Equal("@SQL=\"urn:schemas:httpmail:datereceived\" <= '2025-01-01 00:00'", filter);
    }

    [Fact]
    public void Build_ReceivedAfter_ConvertsOffsetToUtc()
    {
        // Outlook compares urn:schemas:httpmail:datereceived in UTC. Emitting the caller's local
        // wall-clock time makes Restrict drop every message whose UTC timestamp falls inside the
        // offset window - silently, and before the client-side check can recover it. Verified
        // against classic Outlook: a +02:00 caller lost two hours' worth of matching mail.
        var after = new DateTimeOffset(2024, 3, 7, 14, 30, 0, TimeSpan.FromHours(2));

        string? filter = MailRestrictFilter.Build(receivedAfter: after);

        Assert.Equal("@SQL=\"urn:schemas:httpmail:datereceived\" >= '2024-03-07 12:29'", filter);
    }

    [Fact]
    public void Build_ReceivedBefore_ConvertsOffsetToUtc()
    {
        var before = new DateTimeOffset(2024, 3, 7, 14, 30, 0, TimeSpan.FromHours(-5));

        string? filter = MailRestrictFilter.Build(receivedBefore: before);

        Assert.Equal("@SQL=\"urn:schemas:httpmail:datereceived\" <= '2024-03-07 19:31'", filter);
    }

    [Fact]
    public void Build_ReceivedBounds_SlackNeverNarrowsTheWindow()
    {
        // Sub-minute components must never be truncated in a direction that excludes a match.
        var after = new DateTimeOffset(2024, 3, 7, 14, 30, 45, TimeSpan.Zero);
        var before = new DateTimeOffset(2024, 3, 7, 16, 30, 45, TimeSpan.Zero);

        string? filter = MailRestrictFilter.Build(receivedAfter: after, receivedBefore: before);

        Assert.Contains(">= '2024-03-07 14:29'", filter);
        Assert.Contains("<= '2024-03-07 16:31'", filter);
    }

    [Fact]
    public void Build_MultiplePredicates_AreAndedTogether()
    {
        string? filter = MailRestrictFilter.Build(
            unreadOnly: true,
            subjectContains: "budget",
            hasAttachment: true);

        Assert.Equal(
            "@SQL=\"urn:schemas:httpmail:read\" = 0 " +
            "AND \"urn:schemas:httpmail:subject\" LIKE '%budget%' " +
            "AND \"urn:schemas:httpmail:hasattachment\" = 1",
            filter);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Build_BlankStringValues_AreTreatedAsUnset(string value)
    {
        // An empty string arriving from a CLI flag must not become LIKE '%%', which would be a
        // no-op predicate that costs a Restrict call for nothing.
        string? filter = MailRestrictFilter.Build(subjectContains: value, fromAddress: value);

        Assert.Null(filter);
    }

    [Fact]
    public void Build_SingleQuotedPropertyReference_IsNotVulnerableToFilterInjection()
    {
        // A value crafted to close the literal and append a clause must stay inside the literal.
        string? filter = MailRestrictFilter.Build(subjectContains: "x' OR '1'='1");

        Assert.Equal("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%x'' OR ''1''=''1%'", filter);
    }

    // ---- Content-index full-text pushdown (#42) ----

    [Fact]
    public void Build_WithFullTextQuery_AsksTheContentIndexAcrossBodyAndHeaders()
    {
        // The client-side check this replaces matches subject, sender name, sender address, To, Cc
        // *or* body. Pushing only the body down would be under-inclusive - Outlook would drop a
        // subject match before the client ever saw it, and the caller would be told the mail does
        // not exist. So the pushed-down clause has to span the same fields.
        string? filter = MailRestrictFilter.Build(fullTextQuery: "invoice");

        Assert.Equal(
            "@SQL=(\"urn:schemas:httpmail:textdescription\" ci_phrasematch 'invoice' "
            + "OR \"urn:schemas:httpmail:subject\" ci_phrasematch 'invoice' "
            + "OR \"urn:schemas:httpmail:fromname\" ci_phrasematch 'invoice' "
            + "OR \"urn:schemas:httpmail:fromemail\" ci_phrasematch 'invoice' "
            + "OR \"urn:schemas:httpmail:displayto\" ci_phrasematch 'invoice' "
            + "OR \"urn:schemas:httpmail:displaycc\" ci_phrasematch 'invoice')",
            filter);
    }

    [Fact]
    public void Build_WithFullTextQuery_EscapesAnEmbeddedQuote()
    {
        string? filter = MailRestrictFilter.Build(fullTextQuery: "o'brien");

        Assert.NotNull(filter);
        Assert.Contains("ci_phrasematch 'o''brien'", filter!, StringComparison.Ordinal);
        Assert.DoesNotContain("ci_phrasematch 'o'brien'", filter!, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_WithFullTextQuery_AndStructuredPredicates_CombinesThemWithAnd()
    {
        string? filter = MailRestrictFilter.Build(unreadOnly: true, fullTextQuery: "invoice");

        Assert.NotNull(filter);
        Assert.StartsWith("@SQL=\"urn:schemas:httpmail:read\" = 0 AND (", filter!, StringComparison.Ordinal);
        Assert.Contains("ci_phrasematch 'invoice'", filter!, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_WithBlankFullTextQuery_EmitsNoContentIndexClause()
    {
        string? filter = MailRestrictFilter.Build(fullTextQuery: "   ");

        Assert.Null(filter);
    }

    [Fact]
    public void Build_WithFullTextQuery_IsNotVulnerableToFilterInjection()
    {
        string? filter = MailRestrictFilter.Build(fullTextQuery: "x' OR '1'='1");

        Assert.NotNull(filter);
        Assert.Contains("ci_phrasematch 'x'' OR ''1''=''1'", filter!, StringComparison.Ordinal);
    }
}
