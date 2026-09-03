using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Paging behaviour for <c>mail.list</c> / <c>mail.search</c> against a real mailbox (#43).
///
/// <para>
/// These are read-only. They page over a folder the profile already has and compare the result
/// against a single unpaged call, so nothing is created, moved or deleted.
/// </para>
///
/// <para>
/// The property under test is the one a caller actually depends on: <b>walking the pages must yield
/// exactly what one big call yields</b> - every item once, none missing. Asserting only "a cursor
/// came back" would pass against a cursor that silently skipped half the folder, which is precisely
/// the failure this contract exists to prevent. Before #43 there was no cursor at all and a caller
/// could not reach page two, so these cannot pass against the previous behaviour.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailPaging")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailPagingTests(ITestOutputHelper output)
{
    private const string PagedFolder = "drafts";

    /// <summary>
    /// The whole contract in one assertion: paging through a folder returns the same items, in the
    /// same order, as asking for them all at once.
    ///
    /// <para>
    /// Compares a bounded prefix rather than the entire folder so the test does not depend on the
    /// folder fitting inside a single call's cap - an earlier version silently skipped itself on
    /// this machine for exactly that reason, which would have left the central claim unverified
    /// while the suite still reported green.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void MailList_WalkedInSmallPages_ReturnsTheSameItemsAsOneLargeCall()
    {
        EnsureOutlookAvailable();

        const int Prefix = 10;
        var commands = new MailCommands();

        var single = commands.List(folder: PagedFolder, maxCount: Prefix);
        Assert.True(single.Success, single.ErrorMessage);
        Skip.If(single.ReturnedCount < Prefix, $"'{PagedFolder}' holds fewer than {Prefix} mail items.");

        var expected = single.Messages.Select(m => m.EntryId).ToList();
        var walked = WalkPages(commands, pageSize: 2, stopAfter: Prefix, out int pages);

        output.WriteLine($"one call: {expected.Count} items; paged walk: {walked.Count} items over {pages} pages.");

        // Sequence equality, not set equality. Paging is a keyset walk over a stated ordering, so a
        // cursor that returned the right items in the wrong order would still be broken.
        Assert.Equal(expected, walked);
    }

    /// <summary>
    /// A duplicate is the failure an offset-based cursor produces the moment mail arrives mid-walk.
    /// Called out separately from the set comparison because it is the specific regression #43
    /// names.
    /// </summary>
    [SkippableFact]
    public void MailList_WalkedInSmallPages_ReturnsNoItemTwice()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        Skip.If(commands.List(folder: PagedFolder, maxCount: 3).ReturnedCount < 3,
            $"'{PagedFolder}' holds too few mail items to page over.");

        var walked = WalkAllPages(commands, pageSize: 2, out _);

        Assert.Equal(walked.Count, walked.Distinct(StringComparer.Ordinal).Count());
    }

    /// <summary>
    /// The walk has to end on its own. A cursor that failed to advance would keep reporting more
    /// results forever, which is worse than truncation because the caller cannot detect it.
    /// </summary>
    [SkippableFact]
    public void MailList_WalkedInSmallPages_TerminatesWithoutACursor()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var first = commands.List(folder: PagedFolder, maxCount: 2);
        Assert.True(first.Success, first.ErrorMessage);
        Skip.If(!first.HasMore, $"'{PagedFolder}' fits in a single page.");

        MailListResultShape last = WalkToEnd(commands, pageSize: 2);

        Assert.False(last.HasMore);
        Assert.Null(last.NextCursor);
    }

    [SkippableFact]
    public void MailList_FirstPage_ReportsTheOrderingItPagesOver()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().List(folder: PagedFolder, maxCount: 2);

        Assert.True(result.Success, result.ErrorMessage);
        // Paging is a keyset walk over this ordering, so a caller must not have to guess it.
        Assert.Equal("receivedTime", result.SortedBy);
        Assert.Equal("descending", result.SortDirection);
    }

    /// <summary>
    /// A cursor minted by one query must not continue a different one. Honouring it would return a
    /// page of the wrong result set while looking entirely successful.
    /// </summary>
    [SkippableFact]
    public void MailList_WithACursorFromADifferentQuery_FailsInsteadOfReturningTheWrongPage()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var first = commands.List(folder: PagedFolder, maxCount: 1);
        Assert.True(first.Success, first.ErrorMessage);
        Skip.If(first.NextCursor is null, $"'{PagedFolder}' fits in a single page.");

        // Same folder, different filter - so a different result set.
        var reused = commands.List(folder: PagedFolder, maxCount: 1, unreadOnly: true, cursor: first.NextCursor);

        Assert.False(reused.Success);
        Assert.Contains("different query", reused.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A cursor that cannot be read must fail loudly. Silently restarting would re-serve page one,
    /// so a caller looping on <c>hasMore</c> would never terminate.
    /// </summary>
    [SkippableFact]
    public void MailList_WithAMalformedCursor_FailsInsteadOfRestarting()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().List(folder: PagedFolder, maxCount: 2, cursor: "not-a-real-cursor");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Empty(result.Messages);
    }

    /// <summary>
    /// Page size is not part of what a cursor is bound to, so a caller may shrink or grow pages
    /// part-way through a walk. It changes how much comes back, never which result set.
    /// </summary>
    [SkippableFact]
    public void MailList_WithADifferentPageSizeMidWalk_ContinuesTheSameWalk()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var first = commands.List(folder: PagedFolder, maxCount: 1);
        Assert.True(first.Success, first.ErrorMessage);
        Skip.If(first.NextCursor is null, $"'{PagedFolder}' fits in a single page.");

        var second = commands.List(folder: PagedFolder, maxCount: 3, cursor: first.NextCursor);

        Assert.True(second.Success, second.ErrorMessage);
        Assert.DoesNotContain(
            second.Messages,
            m => first.Messages.Any(f => string.Equals(f.EntryId, m.EntryId, StringComparison.Ordinal)));
    }

    private static List<string?> WalkAllPages(MailCommands commands, int pageSize, out int pages)
        => WalkPages(commands, pageSize, stopAfter: int.MaxValue, out pages);

    /// <summary>
    /// Walks pages until the walk ends or <paramref name="stopAfter"/> items have been collected,
    /// returning exactly that many. The page-count ceiling is an assertion in its own right: a
    /// cursor that failed to advance would otherwise loop here forever rather than fail.
    /// </summary>
    private static List<string?> WalkPages(MailCommands commands, int pageSize, int stopAfter, out int pages)
    {
        var collected = new List<string?>();
        string? cursor = null;
        pages = 0;

        while (pages < 200)
        {
            var page = commands.List(folder: PagedFolder, maxCount: pageSize, cursor: cursor);
            Assert.True(page.Success, page.ErrorMessage);
            pages++;
            collected.AddRange(page.Messages.Select(m => m.EntryId));

            if (collected.Count >= stopAfter)
            {
                return [.. collected.Take(stopAfter)];
            }

            if (!page.HasMore)
            {
                return collected;
            }

            Assert.False(string.IsNullOrWhiteSpace(page.NextCursor));
            cursor = page.NextCursor;
        }

        Assert.Fail("Paging did not terminate within 200 pages, so the cursor is not advancing.");
        return collected;
    }

    private static MailListResultShape WalkToEnd(MailCommands commands, int pageSize)
    {
        string? cursor = null;

        for (int pages = 0; pages < 200; pages++)
        {
            var page = commands.List(folder: PagedFolder, maxCount: pageSize, cursor: cursor);
            Assert.True(page.Success, page.ErrorMessage);

            if (!page.HasMore)
            {
                return new MailListResultShape(page.HasMore, page.NextCursor);
            }

            cursor = page.NextCursor;
        }

        Assert.Fail("Paging did not terminate within 200 pages, so the cursor is not advancing.");
        return new MailListResultShape(true, cursor);
    }

    private sealed record MailListResultShape(bool HasMore, string? NextCursor);

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping Outlook paging test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInterop.NameSpace? session = null;

        try
        {
            session = application.GetNamespace("MAPI");
            _ = session.Folders.Count;
        }
        catch (Exception ex)
        {
            output.WriteLine($"Skipping Outlook paging test: {ex.Message}");
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            if (session != null && Marshal.IsComObject(session))
            {
                _ = Marshal.FinalReleaseComObject(session);
            }

            // The shared Application is reused by the runner, so it is released rather than
            // final-released - tearing down its RCW would break every later call.
            if (Marshal.IsComObject(application))
            {
                _ = Marshal.ReleaseComObject(application);
            }
        }
    }
}
