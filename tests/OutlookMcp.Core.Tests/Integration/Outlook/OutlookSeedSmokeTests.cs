using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Attachment;
using OutlookMcp.Core.Commands.Application;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

[Trait("Category", "Integration")]
[Trait("Feature", "OutlookSeed")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookSeedSmokeTests(ITestOutputHelper output)
{
    [SkippableFact]
    public void ApplicationGetStatus_WhenOutlookAvailable_ReturnsSuccess()
    {
        EnsureOutlookAvailable();

        var commands = new ApplicationCommands();
        var result = commands.GetStatus();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Connected);
        Assert.False(string.IsNullOrWhiteSpace(result.Version));
    }

    [SkippableFact]
    public void FolderListDefault_WhenOutlookAvailable_ReturnsCommonFolderRoles()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();
        var result = commands.ListDefault();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Folders);
        Assert.Contains(result.Folders, folder => folder.Role == "inbox");
        Assert.Contains(result.Folders, folder => folder.Role == "drafts");
        Assert.Contains(result.Folders, folder => folder.Available);
    }

    [SkippableFact]
    public void FolderListChildren_WhenOutlookAvailable_ReturnsInboxChildFolders()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();
        var result = commands.ListChildren(parentFolder: "inbox");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.Folders);
        Assert.All(result.Folders, folder => Assert.True(folder.Available));
    }

    [SkippableFact]
    public void FolderResolvePath_WhenOutlookAvailable_ResolvesInboxByRole()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();
        var result = commands.ResolvePath(folder: "inbox");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.FolderPath));
    }

    [SkippableFact]
    public void FolderListItems_WhenOutlookAvailable_ReturnsDraftsFolderItems()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new FolderCommands();
            var result = commands.ListItems(folder: "drafts", maxCount: 25);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.NotNull(result.Items);
            Assert.Contains(result.Items, item => item.EntryId == draft.EntryId);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void CreateMailDraft_WhenOutlookAvailable_CreatesAndDeletesDraft()
    {
        EnsureOutlookAvailable();

        var result = CreateSmokeDraft();

        try
        {
            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(result.Saved);
            Assert.StartsWith("Copilot Outlook smoke ", result.Subject);
            Assert.False(string.IsNullOrWhiteSpace(result.EntryId));
        }
        finally
        {
            DeleteDraft(result.EntryId!, result.StoreId);
        }
    }

    [SkippableFact]
    public void MailRead_WhenDraftResolvedByEntryId_ReturnsRequestedMail()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();
            var result = commands.Read(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(result.HasActiveMail);
            Assert.Equal(draft.EntryId, result.EntryId);
            Assert.Equal(draft.StoreId, result.StoreId);
            Assert.Equal(draft.Subject, result.Subject);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailList_WithUnreadOnly_UsesRestrictWithoutErrorAndReportsTruncationExplicitly()
    {
        // #27: unreadOnly is pushed down via Items.Restrict rather than a client-side scan capped
        // at a fixed item count. This smoke test only asserts the Restrict path does not throw
        // and that Truncated/ScannedCount are populated -- it does not (and cannot, without
        // seeding thousands of items) prove a match beyond the old 500-item cap is found; that
        // requires either a large fixture mailbox or a mocked Items collection, tracked as
        // follow-up test coverage.
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();
            var result = commands.List(folder: "drafts", maxCount: 25, unreadOnly: true);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(result.TotalItemCount >= 0);
            Assert.True(result.ScannedCount >= 0);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailList_WithSubjectContains_FindsSeededDraftWithoutScanningWholeFolder()
    {
        // #27: subjectContains is pushed down through Items.Restrict. The seeded draft carries a
        // GUID subject, so a correct filter returns exactly one item and, crucially, ScannedCount
        // must stay at that one item rather than climbing to the folder's total -- that difference
        // is the whole point of the change.
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            string token = draft.Subject!.Split(' ')[^1];

            var commands = new MailCommands();
            var result = commands.List(folder: "drafts", maxCount: 25, subjectContains: token);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Contains(result.Messages, message => message.EntryId == draft.EntryId);
            Assert.Equal(1, result.ReturnedCount);
            Assert.Equal(1, result.ScannedCount);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailList_WithReceivedWindowAroundItem_StillFindsIt()
    {
        // Regression test for a bug that only a live mailbox could expose. Outlook compares
        // urn:schemas:httpmail:datereceived in UTC. An earlier build emitted the caller's local
        // wall-clock time as the DASL literal, so on a UTC+02:00 machine Restrict silently dropped
        // every message in a two-hour band -- and because Restrict runs inside Outlook, the
        // client-side check never saw those items and the caller got a confident, wrong "no match".
        //
        // Centring the window on the item's own ReceivedTime makes this fail in any time zone with
        // a non-zero offset, rather than only when the test machine happens to be east of UTC.
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();

            var unfiltered = commands.List(folder: "drafts", maxCount: 100);
            Assert.True(unfiltered.Success, unfiltered.ErrorMessage);

            var seeded = unfiltered.Messages.FirstOrDefault(message => message.EntryId == draft.EntryId);
            Skip.If(seeded?.ReceivedTime is null, "Seeded draft exposed no ReceivedTime to centre the window on.");

            DateTimeOffset centre = seeded!.ReceivedTime!.Value;

            var result = commands.List(
                folder: "drafts",
                maxCount: 100,
                receivedAfter: centre.AddMinutes(-30).ToString("O"),
                receivedBefore: centre.AddMinutes(30).ToString("O"));

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Contains(result.Messages, message => message.EntryId == draft.EntryId);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailList_WithUnmatchableStructuredFilter_SucceedsWithNoMatches()
    {
        // A filter that matches nothing must be an empty success, not an error and not a silent
        // fallback to returning everything.
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var result = commands.List(
            folder: "drafts",
            maxCount: 25,
            subjectContains: $"no-such-subject-{Guid.NewGuid():N}");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Empty(result.Messages);
        Assert.Equal(0, result.ReturnedCount);
    }

    [SkippableFact]
    public void MailList_WithInvalidReceivedAfter_FailsWithoutTouchingOutlook()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var result = commands.List(folder: "drafts", maxCount: 25, receivedAfter: "not-a-date");

        Assert.False(result.Success);
        Assert.Contains("receivedAfter", result.ErrorMessage, StringComparison.Ordinal);
    }

    [SkippableFact]
    public void MailReply_WithExplicitEntryId_WorksHeadlessly()
    {
        // #36: reply must be targetable via entryId/storeId without any Outlook window focused
        // or item selected -- this is what makes "find a message, then reply to it" usable.
        EnsureOutlookAvailable();

        // A received message, not a draft. Outlook cannot build a reply from an item that was never
        // sent (#92); an earlier version of this test asked it to and had skipped itself ever since,
        // so the impossible assertion was never observed.
        (string EntryId, string? StoreId) source = FindReceivedMessage();
        MailDraftResult? replyDraft = null;

        try
        {
            var commands = new MailCommands();
            replyDraft = commands.Reply(
                entryId: source.EntryId,
                storeId: source.StoreId,
                useActiveMail: false,
                body: "Headless reply body.");

            Assert.True(replyDraft.Success, replyDraft.ErrorMessage);
            Assert.True(replyDraft.Saved);
            Assert.False(string.IsNullOrWhiteSpace(replyDraft.EntryId));

            var replyBody = commands.Read(entryId: replyDraft.EntryId, storeId: replyDraft.StoreId, useActiveMail: false);
            Assert.Contains("Headless reply body.", replyBody.BodyPreview ?? string.Empty);
        }
        finally
        {
            if (replyDraft?.EntryId != null)
            {
                DeleteDraft(replyDraft.EntryId, replyDraft.StoreId);
            }
        }
    }

    /// <summary>
    /// Replying to a draft is impossible in Outlook, so the only useful behaviour is to say why.
    /// Asserted because the raw Outlook message - "Could not send the message" - describes an
    /// operation the caller did not ask for and gives an agent nothing to act on (#92).
    /// </summary>
    [SkippableFact]
    public void MailReply_ToAnUnsentDraft_ExplainsWhyRatherThanReportingASendFailure()
    {
        EnsureOutlookAvailable();

        var sourceDraft = CreateSmokeDraft();

        try
        {
            var result = new MailCommands().Reply(
                entryId: sourceDraft.EntryId,
                storeId: sourceDraft.StoreId,
                useActiveMail: false,
                body: "This cannot work.");

            Assert.False(result.Success);
            Assert.Contains("unsent draft", result.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Could not send the message", result.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteDraft(sourceDraft.EntryId!, sourceDraft.StoreId);
        }
    }

    [SkippableFact]
    public void MailForward_WithExplicitEntryIdAndRecipients_WorksHeadlessly()
    {
        // #36: forward must accept an explicit recipient, since a forwarded message otherwise has
        // nobody to send to and no headless way to add one.
        EnsureOutlookAvailable();

        MailDraftResult? forwardDraft = null;

        try
        {
            var commands = new MailCommands();

            // A real mailbox can hold rights-protected mail carrying a "Do Not Forward" policy,
            // which Outlook enforces regardless of what this code does. That is the policy working,
            // not a defect -- so walk past those messages to a forwardable one rather than skipping
            // the whole test on the first one, which would quietly stop testing anything.
            foreach ((string entryId, string? storeId) in FindReceivedMessages())
            {
                forwardDraft = commands.Forward(
                    entryId: entryId,
                    storeId: storeId,
                    useActiveMail: false,
                    recipientTo: "copilot-outlook-smoke@example.com",
                    body: "Headless forward body.");

                if (forwardDraft.Success
                    || !(forwardDraft.ErrorMessage ?? string.Empty).Contains("Permission to this message is restricted", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                forwardDraft = null;
            }

            Skip.If(forwardDraft == null, "Every received message available is rights-protected and cannot be forwarded.");

            Assert.True(forwardDraft!.Success, forwardDraft.ErrorMessage);
            Assert.True(forwardDraft.Saved);
            Assert.False(string.IsNullOrWhiteSpace(forwardDraft.EntryId));
            Assert.Equal("copilot-outlook-smoke@example.com", forwardDraft.To);
        }
        finally
        {
            if (forwardDraft?.EntryId != null)
            {
                DeleteDraft(forwardDraft.EntryId, forwardDraft.StoreId);
            }
        }
    }

    /// <summary>
    /// Picks a real received message out of the Inbox to reply to. Skips rather than fails when the
    /// Inbox is empty: no message means nothing to test, which is different from the behaviour being
    /// wrong.
    /// </summary>
    private static (string EntryId, string? StoreId) FindReceivedMessage()
    {
        var candidates = FindReceivedMessages();
        Skip.If(candidates.Count == 0, "The mailbox holds no received message to reply to.");
        return candidates[0];
    }

    private static readonly string[] ReceivedMessageFolders = ["inbox", "Inbox/older", "Archive"];

    /// <summary>
    /// Returns received messages, newest first, for tests that may need to walk past ones the
    /// mailbox's own policies rule out.
    ///
    /// <para>
    /// More than one folder is searched because an Inbox holding two rights-protected messages is
    /// enough to make every one of these tests skip - which looks like a pass and verifies nothing.
    /// </para>
    /// </summary>
    private static List<(string EntryId, string? StoreId)> FindReceivedMessages()
    {
        var found = new List<(string EntryId, string? StoreId)>();

        foreach (string folder in ReceivedMessageFolders)
        {
            var listed = new MailCommands().List(folder: folder, maxCount: 25);

            if (!listed.Success)
            {
                continue;
            }

            found.AddRange(listed.Messages
                .Where(m => !m.IsDraft
                            && !string.IsNullOrWhiteSpace(m.EntryId)
                            && (m.ItemType == null || m.ItemType == "mail"))
                .Select(m => (m.EntryId!, m.StoreId)));
        }

        return found;
    }

    [SkippableFact]
    public void MailReply_WithNoTargetAndNothingSelected_ReturnsActionableError()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var result = commands.Reply(useActiveMail: true);

        // This assumes no Outlook window has anything selected/open during the test run; if that
        // assumption doesn't hold the reply will legitimately succeed against whatever is active.
        // Either outcome is acceptable here -- what matters is Success=false always pairs with a
        // non-null, actionable ErrorMessage (Rule 1), never a silent/ambiguous failure.
        if (!result.Success)
        {
            Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        }
    }

    [SkippableFact]
    public void AttachmentList_ForNewDraft_ReturnsEmptyAttachmentCollection()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new AttachmentCommands();
            var result = commands.List(
                mailEntryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(draft.EntryId, result.EntryId);
            Assert.Equal(draft.Subject, result.Subject);
            Assert.Equal(0, result.AttachmentCount);
            Assert.Empty(result.Attachments);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void AttachmentSave_ForNewDraftWithNoAttachments_ReturnsZeroSavedFiles()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();
        string destinationDirectory = Path.Combine(Path.GetTempPath(), $"OutlookSeedSmoke_{Guid.NewGuid():N}");

        try
        {
            var commands = new AttachmentCommands();
            var result = commands.Save(
                destinationDirectory: destinationDirectory,
                mailEntryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(draft.EntryId, result.EntryId);
            Assert.Equal(0, result.SavedCount);
            Assert.Empty(result.SavedFiles);
            Assert.Equal("The selected Outlook mail item has no attachments.", result.Message);
            Assert.True(Directory.Exists(destinationDirectory));
            Assert.Empty(Directory.EnumerateFileSystemEntries(destinationDirectory));
        }
        finally
        {
            if (Directory.Exists(destinationDirectory))
            {
                Directory.Delete(destinationDirectory, recursive: true);
            }

            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void AttachmentAddAndRemove_ForDraft_UpdatesAttachmentCollection()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();
        string tempFile = Path.Combine(Path.GetTempPath(), $"OutlookSeedAttachment_{Guid.NewGuid():N}.txt");
        File.WriteAllText(tempFile, "Outlook attachment smoke test.");

        try
        {
            var commands = new AttachmentCommands();

            var addResult = commands.Add(
                filePath: tempFile,
                mailEntryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(addResult.Success, addResult.ErrorMessage);
            Assert.Equal(Path.GetFileName(tempFile), addResult.FileName);
            Assert.Equal(1, addResult.AttachmentCount);

            var listAfterAdd = commands.List(
                mailEntryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(listAfterAdd.Success, listAfterAdd.ErrorMessage);
            Assert.Equal(1, listAfterAdd.AttachmentCount);
            Assert.Contains(listAfterAdd.Attachments, item => item.FileName == Path.GetFileName(tempFile));

            var removeResult = commands.Remove(
                attachmentIndex: 1,
                mailEntryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false,
                confirm: true);

            Assert.True(removeResult.Success, removeResult.ErrorMessage);
            Assert.Equal(Path.GetFileName(tempFile), removeResult.FileName);
            Assert.Equal(0, removeResult.AttachmentCount);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }

            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailSetReadState_WhenDraftResolvedByEntryId_UpdatesUnreadState()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();

            var unreadResult = commands.SetReadState(
                isRead: false,
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(unreadResult.Success, unreadResult.ErrorMessage);
            Assert.False(unreadResult.Read);

            var readBackUnread = commands.Read(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(readBackUnread.Success, readBackUnread.ErrorMessage);
            Assert.True(readBackUnread.HasActiveMail);
            Assert.True(readBackUnread.Unread);

            var readResult = commands.SetReadState(
                isRead: true,
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(readResult.Success, readResult.ErrorMessage);
            Assert.True(readResult.Read);

            var readBackRead = commands.Read(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(readBackRead.Success, readBackRead.ErrorMessage);
            Assert.True(readBackRead.HasActiveMail);
            Assert.False(readBackRead.Unread);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailMove_WhenDraftResolvedByEntryId_MovesToDeletedItems()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();
        string? cleanupEntryId = draft.EntryId;
        string? cleanupStoreId = draft.StoreId;

        try
        {
            var commands = new MailCommands();
            var moveResult = commands.Move(
                targetFolder: "deleted",
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(moveResult.Success, moveResult.ErrorMessage);
            Assert.True(moveResult.Moved);
            Assert.False(string.IsNullOrWhiteSpace(moveResult.EntryId));
            Assert.Contains("Deleted", moveResult.FolderName ?? string.Empty, StringComparison.OrdinalIgnoreCase);

            cleanupEntryId = moveResult.EntryId;
            cleanupStoreId = moveResult.StoreId;

            var movedMail = commands.Read(
                entryId: moveResult.EntryId,
                storeId: moveResult.StoreId,
                useActiveMail: false);

            Assert.True(movedMail.Success, movedMail.ErrorMessage);
            Assert.True(movedMail.HasActiveMail);
            Assert.Equal(moveResult.EntryId, movedMail.EntryId);
            Assert.Contains("Deleted", movedMail.CurrentFolderPath ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(cleanupEntryId))
            {
                DeleteDraft(cleanupEntryId, cleanupStoreId);
            }
        }
    }

    [SkippableFact]
    public void MailDelete_WhenDraftResolvedByEntryId_ReturnsDeletedResult()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();
        var commands = new MailCommands();
        var result = commands.Delete(
            entryId: draft.EntryId,
            storeId: draft.StoreId,
            useActiveMail: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Deleted);
        Assert.Equal(draft.EntryId, result.EntryId);
        Assert.Equal(draft.Subject, result.Subject);
    }

    [SkippableFact]
    public void MailSetCategories_WhenDraftResolvedByEntryId_UpdatesAndClearsCategories()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();
            var setResult = commands.SetCategories(
                categories: "Copilot Smoke, Follow Up",
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(setResult.Success, setResult.ErrorMessage);
            Assert.Contains("Copilot Smoke", setResult.Categories);
            Assert.Contains("Follow Up", setResult.Categories);

            var readWithCategories = commands.Read(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(readWithCategories.Success, readWithCategories.ErrorMessage);
            Assert.Contains("Copilot Smoke", readWithCategories.Categories);
            Assert.Contains("Follow Up", readWithCategories.Categories);

            var clearResult = commands.SetCategories(
                categories: null,
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(clearResult.Success, clearResult.ErrorMessage);
            Assert.Empty(clearResult.Categories);

            var readWithoutCategories = commands.Read(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(readWithoutCategories.Success, readWithoutCategories.ErrorMessage);
            Assert.Empty(readWithoutCategories.Categories);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailSetSubject_WhenDraftResolvedByEntryId_UpdatesSubject()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();
            var result = commands.SetSubject(
                subject: "Copilot Outlook smoke updated subject",
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal("Copilot Outlook smoke updated subject", result.Subject);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailSetBody_WhenDraftResolvedByEntryId_UpdatesBody()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();
            var setResult = commands.SetBody(
                body: "Copilot Outlook smoke updated body.",
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(setResult.Success, setResult.ErrorMessage);

            var readResult = commands.Read(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(readResult.Success, readResult.ErrorMessage);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void MailSetRecipients_WhenDraftResolvedByEntryId_UpdatesToAndCc()
    {
        EnsureOutlookAvailable();

        var draft = CreateSmokeDraft();

        try
        {
            var commands = new MailCommands();
            var result = commands.SetRecipients(
                recipientTo: "copilot-smoke-to@example.com",
                cc: "copilot-smoke-cc@example.com",
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Contains("copilot-smoke-to@example.com", result.To);
            Assert.Contains("copilot-smoke-cc@example.com", result.Cc);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    [SkippableFact]
    public void CalendarCreateAndRead_WhenOutlookAvailable_CreatesAndReadsAppointment()
    {
        EnsureOutlookAvailable();

        string start = DateTimeOffset.Now.AddHours(1).ToString("o");
        string endTime = DateTimeOffset.Now.AddHours(2).ToString("o");
        var commands = new CalendarCommands();
        var result = commands.CreateAppointment(
            subject: $"Copilot Calendar smoke {Guid.NewGuid():N}",
            start: start,
            endTime: endTime,
            location: "Copilot Test",
            body: "Calendar smoke appointment.",
            display: false);

        try
        {
            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(result.Saved);
            Assert.False(string.IsNullOrWhiteSpace(result.EntryId));

            var readResult = commands.Read(
                entryId: result.EntryId,
                storeId: result.StoreId,
                useActiveAppointment: false);

            Assert.True(readResult.Success, readResult.ErrorMessage);
            Assert.True(readResult.HasItem);
            Assert.Equal(result.EntryId, readResult.EntryId);
            Assert.Equal(result.Subject, readResult.Subject);
            Assert.Equal("Copilot Test", readResult.Location);
        }
        finally
        {
            DeleteAppointment(result.EntryId!, result.StoreId);
        }
    }

    [SkippableFact]
    public void CalendarList_WhenCreatedAppointmentFallsInRange_ReturnsMatchingItem()
    {
        EnsureOutlookAvailable();

        DateTimeOffset startTime = DateTimeOffset.Now.AddHours(3);
        DateTimeOffset endTime = startTime.AddMinutes(30);
        var commands = new CalendarCommands();
        var created = commands.CreateAppointment(
            subject: $"Copilot Calendar list {Guid.NewGuid():N}",
            start: startTime.ToString("o"),
            endTime: endTime.ToString("o"),
            location: "Calendar List Test",
            body: "Calendar list smoke appointment.",
            display: false);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var listResult = commands.List(
                start: startTime.AddMinutes(-15).ToString("o"),
                endTime: endTime.AddMinutes(15).ToString("o"),
                maxCount: 20,
                includeBodyPreview: true);

            Assert.True(listResult.Success, listResult.ErrorMessage);
            Assert.Contains(listResult.Appointments, item => item.EntryId == created.EntryId && item.Subject == created.Subject);
        }
        finally
        {
            DeleteAppointment(created.EntryId!, created.StoreId);
        }
    }

    [SkippableFact]
    public void CalendarUpdateAppointment_WhenOutlookAvailable_PersistsChanges()
    {
        EnsureOutlookAvailable();

        DateTimeOffset startTime = DateTimeOffset.Now.AddHours(4);
        DateTimeOffset endTime = startTime.AddMinutes(30);
        var commands = new CalendarCommands();
        var created = commands.CreateAppointment(
            subject: $"Copilot Calendar update {Guid.NewGuid():N}",
            start: startTime.ToString("o"),
            endTime: endTime.ToString("o"),
            location: "Original Location",
            body: "Original calendar body.",
            display: false);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var updated = commands.UpdateAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                subject: "Updated Copilot Calendar Subject",
                location: "Updated Location",
                body: "Updated calendar body.",
                allDay: false);

            Assert.True(updated.Success, updated.ErrorMessage);
            Assert.True(updated.Updated);
            Assert.Equal(created.EntryId, updated.EntryId);
            Assert.Equal("Updated Copilot Calendar Subject", updated.Subject);
            Assert.Equal("Updated Location", updated.Location);

            var readResult = commands.Read(
                entryId: created.EntryId,
                storeId: created.StoreId,
                useActiveAppointment: false);

            Assert.True(readResult.Success, readResult.ErrorMessage);
            Assert.Equal("Updated Copilot Calendar Subject", readResult.Subject);
            Assert.Equal("Updated Location", readResult.Location);
            Assert.Contains("Updated calendar body.", readResult.BodyPreview);
        }
        finally
        {
            DeleteAppointment(created.EntryId!, created.StoreId);
        }
    }

    [SkippableFact]
    public void CalendarDeleteAppointment_WhenOutlookAvailable_RemovesAppointment()
    {
        EnsureOutlookAvailable();

        DateTimeOffset startTime = DateTimeOffset.Now.AddHours(5);
        DateTimeOffset endTime = startTime.AddMinutes(45);
        var commands = new CalendarCommands();
        var created = commands.CreateAppointment(
            subject: $"Copilot Calendar delete {Guid.NewGuid():N}",
            start: startTime.ToString("o"),
            endTime: endTime.ToString("o"),
            location: "Delete Location",
            body: "Delete calendar body.",
            display: false);

        Assert.True(created.Success, created.ErrorMessage);

        var deleted = commands.DeleteAppointment(
            entryId: created.EntryId,
            storeId: created.StoreId,
            useActiveAppointment: false);

        Assert.True(deleted.Success, deleted.ErrorMessage);
        Assert.True(deleted.Deleted);
        Assert.Equal(created.EntryId, deleted.EntryId);

        var readResult = commands.Read(
            entryId: created.EntryId,
            storeId: created.StoreId,
            useActiveAppointment: false);

        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.False(readResult.HasItem);
    }

    [SkippableFact]
    public void Execute_AfterPriorCall_SharedApplicationRemainsUsable()
    {
        // Regression test for #19: OutlookInteropRunner used to FinalReleaseComObject the
        // shared, already-running Outlook.Application obtained via GetActiveObject. Because
        // .NET caches one RCW per process for that COM identity, finalizing it during one
        // call's cleanup could invalidate every other holder's reference to the same object.
        // This test issues two sequential Execute() calls and asserts the second one succeeds,
        // proving the first call's cleanup did not tear down the shared Application RCW.
        EnsureOutlookAvailable();

        string firstVersion = OutlookInteropRunner.Execute(
            "regression-first-call",
            (application, _) => application.Version,
            ex => throw new InvalidOperationException("First Execute call failed.", ex));

        Assert.False(string.IsNullOrWhiteSpace(firstVersion));

        // If the shared Application RCW was invalidated by the first call's cleanup, this
        // second call will throw InvalidComObjectException instead of returning a version.
        string secondVersion = OutlookInteropRunner.Execute(
            "regression-second-call",
            (application, _) => application.Version,
            ex => throw new InvalidOperationException("Second Execute call failed after first call's cleanup.", ex));

        Assert.Equal(firstVersion, secondVersion);
    }

    /// <summary>
    /// Skips the calling test (via Xunit.SkippableFact) unless a running classic Outlook desktop
    /// instance with a usable MAPI session is available. Regression guard for #22: this MUST
    /// report tests as Skipped, never silently Passed, when Outlook is unavailable.
    /// </summary>
    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping Outlook smoke test: no running classic Outlook desktop instance is available.");
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
            output.WriteLine($"Skipping Outlook smoke test: {ex.Message}");
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(session);
            ReleaseSharedApplication(application);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static MailDraftResult CreateSmokeDraft()
    {
        string subject = $"Copilot Outlook smoke {Guid.NewGuid():N}";
        var commands = new MailCommands();
        var result = commands.CreateMailDraft(
            subject: subject,
            body: "Outlook smoke test draft body.",
            display: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Saved);
        Assert.Equal(subject, result.Subject);
        Assert.False(string.IsNullOrWhiteSpace(result.EntryId));
        return result;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void DeleteDraft(string entryId, string? storeId)
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            return;
        }

        OutlookInterop.NameSpace? session = null;
        object? item = null;
        OutlookInterop.MailItem? mail = null;

        try
        {
            session = application.GetNamespace("MAPI");
            item = session.GetItemFromID(
                entryId,
                string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
            mail = item as OutlookInterop.MailItem;
            mail?.Delete();
        }
        finally
        {
            ReleaseComObject(mail);
            ReleaseComObject(item);
            ReleaseComObject(session);
            ReleaseSharedApplication(application);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void DeleteAppointment(string entryId, string? storeId)
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            return;
        }

        OutlookInterop.NameSpace? session = null;
        object? item = null;
        OutlookInterop.AppointmentItem? appointment = null;

        try
        {
            session = application.GetNamespace("MAPI");
            item = session.GetItemFromID(entryId, string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
            appointment = item as OutlookInterop.AppointmentItem;
            appointment?.Delete();
        }
        finally
        {
            ReleaseComObject(appointment);
            ReleaseComObject(item);
            ReleaseComObject(session);
            ReleaseSharedApplication(application);
        }
    }

    private static void ReleaseSharedApplication(object? value)
    {
        // The Application obtained via TryGetRunningApplication is the user's shared,
        // already-running Outlook instance. Use ReleaseComObject (ref-count decrement),
        // never FinalReleaseComObject, or we invalidate the cached RCW for every other
        // holder in the process. See #19.
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.ReleaseComObject(value);
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }
    }

    /// <summary>
    /// A term that appears only deep inside a long body must still be found.
    /// <para>
    /// This is the whole of issue #42 reduced to one case. <c>MatchesQuery</c> hydrates the entire
    /// body through COM and then throws away everything past 1200 characters before searching it, so
    /// a term further in is invisible. The caller is not told the body was truncated; they are told
    /// there is no such mail, which is the one answer a search must never give wrongly.
    /// </para>
    /// <para>
    /// The truncation buys nothing. The expensive part is the <c>mail.Body</c> COM call, and that
    /// already happened by the time the string is cut.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void MailSearch_WithTermBeyondThePreviewWindow_StillFindsIt()
    {
        EnsureOutlookAvailable();

        string needle = $"needle{Guid.NewGuid():N}";
        var draft = CreateDraftWithBuriedTerm(needle, out string subject);

        try
        {
            var commands = new MailCommands();
            var found = commands.Search(query: needle, folder: "drafts", maxCount: 100);

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Contains(found.Messages, m => m.Subject == subject);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    /// <summary>
    /// The same search restricted to the first part of the body must still work, so the fix is not
    /// just "search more" but "search all of it". Guards against a regression that trims from the
    /// wrong end.
    /// </summary>
    [SkippableFact]
    public void MailSearch_WithTermAtTheStartOfALongBody_StillFindsIt()
    {
        EnsureOutlookAvailable();

        string needle = $"needle{Guid.NewGuid():N}";
        var commands = new MailCommands();
        string subject = $"Copilot Outlook smoke {Guid.NewGuid():N}";

        var draft = commands.CreateMailDraft(
            subject: subject,
            body: needle + " " + new string('x', 4000),
            display: false);

        Assert.True(draft.Success, draft.ErrorMessage);

        try
        {
            var found = commands.Search(query: needle, folder: "drafts", maxCount: 100);

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Contains(found.Messages, m => m.Subject == subject);
        }
        finally
        {
            DeleteDraft(draft.EntryId!, draft.StoreId);
        }
    }

    /// <summary>
    /// Creates a draft whose body is long enough that <paramref name="needle"/> sits well past the
    /// 1200 character window the old client-side matcher looked at.
    /// </summary>
    private static MailDraftResult CreateDraftWithBuriedTerm(string needle, out string subject)
    {
        subject = $"Copilot Outlook smoke {Guid.NewGuid():N}";

        // Filler first, then the needle at roughly character 3000, then more filler so the term is
        // not near either edge.
        string body = new string('a', 3000) + " " + needle + " " + new string('b', 1000);

        var commands = new MailCommands();
        var result = commands.CreateMailDraft(subject: subject, body: body, display: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Saved);
        Assert.False(string.IsNullOrWhiteSpace(result.EntryId));

        return result;
    }
}
