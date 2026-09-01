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
                useActiveMail: false);

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
}
