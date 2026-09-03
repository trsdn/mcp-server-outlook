using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Full-text search via Outlook's content index (#42).
///
/// <para>
/// The default search path hydrates each candidate message and does a substring check on its body.
/// That is exact, but it is bounded: the scan stops at a safety limit, so in a folder larger than
/// that limit a genuine match further back is never seen, and the caller is told - with
/// <c>success: true</c> - that no such mail exists. The content index has no such horizon, because
/// Outlook narrows the folder before the scan begins.
/// </para>
///
/// <para>
/// The two paths do not have identical semantics, and pretending otherwise would be the real danger:
/// a substring check matches <c>foo</c> inside <c>foobar</c>, and a word-based index does not. So the
/// engine is opt-in and, more importantly, the response always names which one answered. These tests
/// exist mostly to hold that reporting honest.
/// </para>
///
/// <para>
/// Everything here is read-only. Nothing is created, moved, sent or deleted.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailSearch")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookFullTextSearchTests(ITestOutputHelper output)
{
    /// <summary>
    /// The default path must keep saying it is the default path. If this ever reports the index, a
    /// caller reasoning about word-vs-substring semantics is being misled.
    /// </summary>
    [SkippableFact]
    public void Search_ByDefault_SaysTheClientScanAnswered()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().Search(query: "the", maxCount: 3);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("clientScan", result.SearchEngine);

        output.WriteLine($"Default search returned {result.ReturnedCount} of {result.ScannedCount} scanned.");
    }

    /// <summary>
    /// The load-bearing test. A term taken from a message that is genuinely in the mailbox must be
    /// found by the index path. Taking the term from a real message rather than hard-coding one is
    /// what makes this deterministic on somebody else's mailbox.
    /// </summary>
    [SkippableFact]
    public void Search_InFullTextMode_FindsATermTakenFromARealMessage()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();

        var listed = commands.List(maxCount: 25);
        Assert.True(listed.Success, listed.ErrorMessage);

        string? term = listed.Messages
            .Select(m => m.Subject)
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .SelectMany(s => s!.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            .Select(w => w.Trim())
            .FirstOrDefault(w => w.Length >= 5 && w.All(char.IsLetter));

        Skip.If(term == null, "No message in the inbox has a subject word usable as a search term.");

        output.WriteLine($"Searching the content index for '{term}'.");

        var result = commands.Search(query: term!, maxCount: 10, searchMode: "fullText");

        Assert.True(result.Success, result.ErrorMessage);

        // If the store has no content index, the tool must say so rather than pretend - and in that
        // case this mailbox cannot demonstrate the index path at all.
        Skip.If(
            result.SearchEngine != "contentIndex",
            $"This store did not answer from the content index ({result.SearchEngine}): {result.Message}");

        Assert.NotEmpty(result.Messages);
        Assert.Contains(
            result.Messages,
            m => m.Subject?.Contains(term!, StringComparison.OrdinalIgnoreCase) == true);

        output.WriteLine($"Index returned {result.ReturnedCount} match(es) after scanning {result.ScannedCount} item(s).");
    }

    /// <summary>
    /// A term nothing can match must come back empty *and* say which engine reached that conclusion.
    /// An empty result whose provenance is unknown is exactly the confidently-wrong answer this whole
    /// area of the project exists to remove.
    /// </summary>
    [SkippableFact]
    public void Search_InFullTextMode_ForATermNothingMatches_ReturnsEmptyAndNamesTheEngine()
    {
        EnsureOutlookAvailable();

        string nonsense = $"zzq{Guid.NewGuid():N}";

        var result = new MailCommands().Search(query: nonsense, maxCount: 10, searchMode: "fullText");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Empty(result.Messages);
        Assert.NotNull(result.SearchEngine);

        output.WriteLine($"'{nonsense}' matched nothing; answered by {result.SearchEngine}.");
    }

    /// <summary>
    /// The index narrows the folder before anything is hydrated. The evidence is a direct comparison:
    /// the same query nothing can match forces the client-side scan to walk the entire folder, while
    /// the index path examines nothing at all. This is the same evidence used for the structured
    /// filters in #84, expressed as a comparison so it does not depend on the mailbox being large.
    /// </summary>
    [SkippableFact]
    public void Search_InFullTextMode_ExaminesFarFewerItemsThanTheClientScan()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string nonsense = $"zzq{Guid.NewGuid():N}";

        var scanned = commands.Search(query: nonsense, maxCount: 10);
        var indexed = commands.Search(query: nonsense, maxCount: 10, searchMode: "fullText");

        Assert.True(scanned.Success, scanned.ErrorMessage);
        Assert.True(indexed.Success, indexed.ErrorMessage);

        Skip.If(
            indexed.SearchEngine != "contentIndex",
            $"This store did not answer from the content index ({indexed.SearchEngine}): {indexed.Message}");

        Skip.If(scanned.TotalItemCount < 5, "The folder is empty enough that neither path has to do anything.");

        output.WriteLine(
            $"clientScan examined {scanned.ScannedCount} item(s); "
            + $"contentIndex examined {indexed.ScannedCount} of {indexed.TotalItemCount}.");

        // Both must agree there is nothing there - otherwise this is comparing two different answers.
        Assert.Empty(scanned.Messages);
        Assert.Empty(indexed.Messages);

        Assert.True(
            indexed.ScannedCount < scanned.ScannedCount,
            $"The index examined {indexed.ScannedCount} items and the scan examined {scanned.ScannedCount}. "
            + "The query was not pushed down.");
    }

    /// <summary>
    /// An unrecognised mode must be refused. Falling back to the default and reporting success would
    /// hand the caller substring semantics while they believed they had asked for the index.
    /// </summary>
    [SkippableFact]
    public void Search_WithAnUnknownSearchMode_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().Search(query: "anything", searchMode: "telepathy");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("searchMode", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
