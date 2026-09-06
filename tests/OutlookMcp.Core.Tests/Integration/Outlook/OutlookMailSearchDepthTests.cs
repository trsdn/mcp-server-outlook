using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// The claim the whole search-depth issue exists to prove (#13, #27): a match that sits <b>past</b>
/// the roughly 500-item window the old client-side scan could reach is actually found, rather than
/// being reported as "no such mail".
///
/// <para>
/// Every other test in this area checks a mechanism - which engine ran, which columns came back.
/// This one checks the outcome the mechanisms exist for, and it is the only test here whose failure
/// would mean the issue was never fixed.
/// </para>
///
/// <para>
/// <b>It must not be able to pass vacuously.</b> On a mailbox whose folders are all smaller than the
/// old window there is nothing to prove, and a test that quietly went green there would be worse than
/// no test: it would stand as evidence for a claim it never examined. So it establishes the depth
/// first and <b>skips with the measured numbers</b> if the mailbox cannot support the experiment -
/// never passes. The skip message names the folder and the item count, so a green suite that skipped
/// this test says so out loud.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailSearchDepth")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailSearchDepthTests(ITestOutputHelper output)
{
    /// <summary>
    /// The scan window the old client-side search stopped at. A match beyond this was invisible.
    /// </summary>
    private const int FormerScanWindow = 500;

    /// <summary>
    /// How far past the old window the target message must sit. Far enough that the result cannot be
    /// explained by an off-by-a-few in how items are counted.
    /// </summary>
    private const int RequiredDepth = 520;

    /// <summary>Folders worth trying, largest first once measured.</summary>
    private static readonly string[] CandidateFolders = ["deleted", "inbox", "sent", "junk", "drafts"];

    [SkippableFact]
    public void MailSearch_FindsAMatchFarPastTheFormerScanWindow()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();

        (string folder, int itemCount) = FindLargestFolder(commands);
        output.WriteLine($"deepest available folder: '{folder}' with {itemCount} items");

        Skip.If(
            itemCount < RequiredDepth,
            $"No folder in this mailbox holds {RequiredDepth} items - the largest is '{folder}' with {itemCount}. "
            + $"The claim under test is that a match past the former ~{FormerScanWindow}-item scan window is still "
            + "found, and this mailbox has no such position to place one at, so the experiment cannot be run. "
            + "This is a genuinely untested criterion here, not a passing one.");

        DeepMessage? target = FindMessagePastTheWindow(commands, folder, out int walked);
        output.WriteLine($"walked {walked} messages in '{folder}'");

        Skip.If(
            target == null,
            $"Walked {walked} messages in '{folder}' without reaching position {RequiredDepth} with a subject "
            + "distinctive enough to search for. The folder reports "
            + $"{itemCount} items, but a listing counts only the messages it can model, so the mailbox cannot "
            + "support the experiment. Untested, not passed.");

        output.WriteLine(
            $"target at position {target!.Position}: received {target.ReceivedTime:O}, "
            + $"searching for '{target.Token}' from subject '{target.Subject}'");

        // The search is the point: a free-text query, no structured filter narrowing it, over a
        // folder where the target sits far past where the old scan gave up.
        List<MailSummaryInfo> matches = SearchAllPages(commands, folder, target.Token, out string? engine);
        output.WriteLine($"engine={engine} matched {matches.Count} messages");

        Assert.Contains(matches, m => string.Equals(m.EntryId, target.EntryId, StringComparison.Ordinal));

        // And the depth was real, not an artefact of the folder having been re-ordered mid-test.
        Assert.True(
            target.Position > FormerScanWindow,
            $"The target sat at position {target.Position}, which is inside the former {FormerScanWindow}-item "
            + "window, so finding it proves nothing.");
    }

    /// <summary>
    /// The same claim on the bounded engine, recorded rather than asserted.
    ///
    /// <para>
    /// <c>clientScan</c> opens each candidate and stops at a safety limit. Whether it reaches this
    /// particular depth is a property of the machine, so demanding a result either way would make the
    /// suite flaky. What is worth pinning down is that if it does <b>not</b> reach it, the response
    /// says so through <c>truncated</c> instead of returning an empty list that reads as "no such
    /// mail" - which was the original bug.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void MailSearch_WithTheBoundedEngine_NeverReportsAMissAsAnEmptyResult()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();

        (string folder, int itemCount) = FindLargestFolder(commands);
        Skip.If(
            itemCount < RequiredDepth,
            $"No folder in this mailbox holds {RequiredDepth} items - the largest is '{folder}' with {itemCount}.");

        DeepMessage? target = FindMessagePastTheWindow(commands, folder, out _);
        Skip.If(target == null, $"Could not reach position {RequiredDepth} in '{folder}'.");

        var scanned = commands.Search(
            query: target!.Token,
            folder: folder,
            maxCount: 100,
            searchMode: "clientScan");

        Assert.True(scanned.Success, scanned.ErrorMessage);
        output.WriteLine(
            $"clientScan: returned={scanned.ReturnedCount} scanned={scanned.ScannedCount} "
            + $"truncated={scanned.Truncated} hasMore={scanned.HasMore}");

        bool found = scanned.Messages.Any(m => string.Equals(m.EntryId, target.EntryId, StringComparison.Ordinal));

        // Either it reached the match, or it admits it stopped early. What it must never do is come
        // back complete and empty, because a caller reads that as proof the mail does not exist.
        Assert.True(
            found || scanned.Truncated || scanned.HasMore,
            "The bounded scan neither found the match nor reported that it stopped early, so an agent would "
            + "have concluded the message does not exist.");
    }

    private sealed record DeepMessage(string EntryId, string Subject, string Token, DateTimeOffset? ReceivedTime, int Position);

    /// <summary>
    /// Picks the folder with the most items, so the test runs against whatever this mailbox can
    /// actually offer rather than assuming a particular folder is large.
    /// </summary>
    private (string Folder, int ItemCount) FindLargestFolder(MailCommands commands)
    {
        string best = CandidateFolders[0];
        int bestCount = -1;

        foreach (string candidate in CandidateFolders)
        {
            var probe = commands.List(folder: candidate, maxCount: 1);
            if (!probe.Success)
            {
                output.WriteLine($"'{candidate}' unavailable: {probe.ErrorMessage}");
                continue;
            }

            output.WriteLine($"'{candidate}' holds {probe.TotalItemCount} items");

            if (probe.TotalItemCount > bestCount)
            {
                bestCount = probe.TotalItemCount;
                best = candidate;
            }
        }

        return (best, Math.Max(bestCount, 0));
    }

    /// <summary>
    /// Pages to a message sitting past <see cref="RequiredDepth"/> whose subject carries a token
    /// distinctive enough to search for.
    ///
    /// <para>
    /// Paging is how depth is established rather than assumed: the position is counted from the
    /// walk, so the assertion later can state where the message actually was.
    /// </para>
    /// </summary>
    private DeepMessage? FindMessagePastTheWindow(MailCommands commands, string folder, out int walked)
    {
        string? cursor = null;
        int position = 0;
        walked = 0;
        output.WriteLine($"walking '{folder}' to position {RequiredDepth}");

        for (int pages = 0; pages < 100; pages++)
        {
            var page = commands.List(folder: folder, maxCount: 100, cursor: cursor);
            Assert.True(page.Success, page.ErrorMessage);

            foreach (MailSummaryInfo message in page.Messages)
            {
                position++;
                walked = position;

                if (position < RequiredDepth)
                {
                    continue;
                }

                string? token = PickSearchableToken(message.Subject);
                if (token != null && !string.IsNullOrWhiteSpace(message.EntryId))
                {
                    return new DeepMessage(message.EntryId!, message.Subject!, token, message.ReceivedTime, position);
                }
            }

            if (!page.HasMore)
            {
                return null;
            }

            cursor = page.NextCursor;
        }

        return null;
    }

    /// <summary>
    /// Picks the longest word in a subject, as a term unlikely to be a stop word and long enough that
    /// a match is meaningful. Returns null when the subject offers nothing usable - a wildcard
    /// character would change which engine answers, and a short token would match half the folder.
    /// </summary>
    private static string? PickSearchableToken(string? subject)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            return null;
        }

        string? best = subject
            .Split([' ', '\t', '\r', '\n', ':', ';', ',', '.', '(', ')', '[', ']', '"', '\'', '/', '\\'],
                StringSplitOptions.RemoveEmptyEntries)
            .Where(word => word.Length is >= 8 and <= 30)
            .Where(word => word.All(char.IsLetterOrDigit))
            .OrderByDescending(word => word.Length)
            .FirstOrDefault();

        return best;
    }

    /// <summary>
    /// Walks every page of a search, so a match that exists but sits on page three cannot be reported
    /// as a miss by a test that only ever looked at page one.
    /// </summary>
    private static List<MailSummaryInfo> SearchAllPages(
        MailCommands commands,
        string folder,
        string token,
        out string? engine)
    {
        var collected = new List<MailSummaryInfo>();
        string? cursor = null;
        engine = null;

        for (int pages = 0; pages < 50; pages++)
        {
            var page = commands.Search(query: token, folder: folder, maxCount: 100, cursor: cursor);
            Assert.True(page.Success, page.ErrorMessage);

            engine ??= page.SearchEngine;
            collected.AddRange(page.Messages);

            if (!page.HasMore || string.IsNullOrWhiteSpace(page.NextCursor))
            {
                return collected;
            }

            cursor = page.NextCursor;
        }

        return collected;
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping Outlook search-depth test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
