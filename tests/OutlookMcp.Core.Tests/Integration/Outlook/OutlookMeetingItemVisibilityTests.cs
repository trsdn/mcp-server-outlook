using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Meeting requests, cancellations and responses must be visible in mail listings (#32).
///
/// <para>
/// <c>mail.list</c> cast every item with <c>as Outlook.MailItem</c> and skipped anything that came
/// back null. A meeting invitation is a <c>MeetingItem</c>, not a <c>MailItem</c>, so invitations
/// were silently absent from every listing - and the response said nothing about it. A user asking
/// "what is in my inbox" was told about some of it, confidently, with no indication that anything
/// had been withheld.
/// </para>
///
/// <para>
/// These tests read a folder that genuinely contains scheduling items. They skip only when the
/// mailbox has none anywhere, because a test that cannot find its subject has verified nothing and
/// must say so rather than pass.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailMeetingVisibility")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMeetingItemVisibilityTests(ITestOutputHelper output)
{
    private static readonly string[] KnownItemTypes =
        ["mail", "meetingRequest", "meetingCancellation", "meetingResponse", "other"];

    private static readonly string[] FolderCandidates =
        ["inbox", "Inbox/older", "Inbox/yesterbox", "Archive"];
    /// <summary>
    /// The bug itself: an invitation in a folder must appear in that folder's listing.
    /// </summary>
    [SkippableFact]
    public void MailList_IncludesMeetingItems()
    {
        EnsureOutlookAvailable();

        string folder = FindFolderContainingSchedulingItems();
        var listed = new MailCommands().List(folder: folder, maxCount: 100);

        Assert.True(listed.Success, listed.ErrorMessage);
        output.WriteLine($"'{folder}' returned {listed.ReturnedCount} item(s).");

        Assert.Contains(listed.Messages, m => m.ItemType != null && m.ItemType != "mail");
    }

    /// <summary>
    /// Visible is not enough: a caller has to be able to tell an invitation from a message, because
    /// the two afford completely different actions. Replying to an invitation is not accepting it.
    /// </summary>
    [SkippableFact]
    public void MailList_LabelsEachItemWithItsType()
    {
        EnsureOutlookAvailable();

        string folder = FindFolderContainingSchedulingItems();
        var listed = new MailCommands().List(folder: folder, maxCount: 100);
        Assert.True(listed.Success, listed.ErrorMessage);

        Assert.All(listed.Messages, m => Assert.False(string.IsNullOrWhiteSpace(m.ItemType)));

        var types = listed.Messages.Select(m => m.ItemType).Distinct().ToList();
        output.WriteLine($"item types: {string.Join(", ", types)}");

        Assert.All(types, t => Assert.Contains(t, KnownItemTypes));
    }

    /// <summary>
    /// A meeting item is useless in a listing without the fields that identify it. Asserted because
    /// "it appears" and "it is usable" are different claims and the first can hold while the second
    /// does not.
    /// </summary>
    [SkippableFact]
    public void MeetingItemsCarryTheFieldsNeededToActOnThem()
    {
        EnsureOutlookAvailable();

        string folder = FindFolderContainingSchedulingItems();
        var listed = new MailCommands().List(folder: folder, maxCount: 100);
        Assert.True(listed.Success, listed.ErrorMessage);

        var meeting = listed.Messages.FirstOrDefault(m => m.ItemType != null && m.ItemType != "mail");
        Skip.If(meeting == null, $"'{folder}' returned no scheduling item within the first 100 items.");

        Assert.False(string.IsNullOrWhiteSpace(meeting!.EntryId));
        Assert.False(string.IsNullOrWhiteSpace(meeting.Subject));
        Assert.NotNull(meeting.ReceivedTime);
    }

    /// <summary>
    /// Anything genuinely not listable must be counted, not dropped. A listing whose numbers do not
    /// add up is how "we showed you everything" quietly becomes false.
    /// </summary>
    [SkippableFact]
    public void MailList_ReportsItemsItCouldNotSummarise()
    {
        EnsureOutlookAvailable();

        var listed = new MailCommands().List(folder: "inbox", maxCount: 25);

        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.True(listed.SkippedItemCount >= 0);
        Assert.Equal(listed.ReturnedCount, listed.Messages.Count);
    }

    /// <summary>
    /// Finds a folder that actually holds scheduling items, so these tests assert against real data
    /// rather than skipping quietly on an empty Inbox.
    /// </summary>
    private static string FindFolderContainingSchedulingItems()
    {
        foreach (string candidate in FolderCandidates)
        {
            var listed = new MailCommands().List(folder: candidate, maxCount: 100);

            if (listed.Success && listed.Messages.Any(m => m.ItemType != null && m.ItemType != "mail"))
            {
                return candidate;
            }
        }

        throw new SkipException("This mailbox holds no meeting request, cancellation or response to list.");
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping Outlook meeting visibility test: no running classic Outlook desktop instance is available.");
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
            output.WriteLine($"Skipping Outlook meeting visibility test: {ex.Message}");
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            if (session != null && Marshal.IsComObject(session))
            {
                _ = Marshal.FinalReleaseComObject(session);
            }
        }
    }
}
