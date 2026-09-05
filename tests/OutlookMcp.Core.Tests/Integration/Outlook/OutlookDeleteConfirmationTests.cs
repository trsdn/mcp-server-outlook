using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// The confirmation gates that cannot be decided from the caller's arguments alone (#9).
///
/// <para>
/// <c>mail.delete</c>, <c>contact.delete</c> and <c>task.delete</c> are deliberately <b>not</b>
/// gated in the ordinary case, because Outlook moves the item to Deleted Items and the user can
/// get it back. That rationale stops holding the moment the item is already <i>in</i> Deleted
/// Items: the second delete is a permanent one and leaves nothing to restore. Whether an item is
/// in Deleted Items is a fact about the running mailbox, so this can only be tested against a real
/// Outlook.
/// </para>
///
/// <para>
/// Everything here is a draft or a folder these tests created moments earlier, named with a GUID
/// under a fixed marker, and swept in a <c>finally</c>. Nothing belonging to the user is touched.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "ConfirmationGate")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookDeleteConfirmationTests(ITestOutputHelper output)
{
    private const string Marker = "OutlookMcp confirm gate test";
    private const string ScratchFolderPrefix = "mcp-confirm-gate-";

    /// <summary>
    /// The full soft-delete-then-permanent-delete lifecycle.
    ///
    /// <para>
    /// The first delete needs no confirmation and must not ask for one - gating a recoverable
    /// action would be ceremony without safety. The second must be refused, because by then the
    /// item is in Deleted Items and deleting it again destroys it.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void MailDelete_OfAnItemAlreadyInDeletedItems_RequiresConfirm()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string subject = $"{Marker} {Guid.NewGuid():N}";
        string? deletedItemsId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: subject, body: "placeholder");
            Assert.True(draft.Success, draft.ErrorMessage);
            Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));

            // First delete: an ordinary soft delete, ungated by design.
            var soft = commands.Delete(entryId: draft.EntryId, useActiveMail: false);
            Assert.True(soft.Success, soft.ErrorMessage);
            Assert.Null(soft.ErrorMessage);

            // Moving between folders reassigns the entry id, so the item has to be found again.
            deletedItemsId = FindInDeletedItems(commands, subject);
            Assert.NotNull(deletedItemsId);

            // Second delete: permanent, and therefore refused without confirmation.
            var refused = commands.Delete(entryId: deletedItemsId, useActiveMail: false, confirm: false);

            output.WriteLine($"Refused as expected: {refused.ErrorMessage}");

            Assert.False(refused.Success);
            Assert.False(refused.Deleted);
            Assert.NotNull(refused.ErrorMessage);
            Assert.Contains("confirm=true", refused.ErrorMessage!, StringComparison.Ordinal);

            // The refusal has to be a refusal, not a delete that reported one: the item is still there.
            Assert.NotNull(FindInDeletedItems(commands, subject));

            // With confirmation it goes, and stays gone.
            var permanent = commands.Delete(entryId: deletedItemsId, useActiveMail: false, confirm: true);
            Assert.True(permanent.Success, permanent.ErrorMessage);
            Assert.Null(permanent.ErrorMessage);

            Assert.Null(FindInDeletedItems(commands, subject));
            deletedItemsId = null;
        }
        finally
        {
            if (deletedItemsId != null)
            {
                _ = commands.Delete(entryId: deletedItemsId, useActiveMail: false, confirm: true);
            }
        }
    }

    /// <summary>
    /// Deleting a folder takes every message and every subfolder in it, so it is gated
    /// unconditionally. The folder must still be there afterwards - a gate that refused and deleted
    /// anyway would pass an assertion made only on the return value.
    /// </summary>
    [SkippableFact]
    public void FolderDelete_WithoutConfirm_LeavesTheFolderInPlace()
    {
        var commands = new FolderCommands();
        string parent = ResolveScratchParent(commands);
        string name = $"{ScratchFolderPrefix}{Guid.NewGuid():N}";

        string? created = null;
        try
        {
            var create = commands.Create(parent, name);
            Assert.True(create.Success, create.ErrorMessage);
            created = create.FolderPath;
            Assert.NotNull(created);

            var refused = commands.Delete(created, confirm: false);

            output.WriteLine($"Refused as expected: {refused.ErrorMessage}");

            Assert.False(refused.Success);
            Assert.NotNull(refused.ErrorMessage);
            Assert.Contains("confirm=true", refused.ErrorMessage!, StringComparison.Ordinal);

            Assert.Contains(
                commands.ListChildren(parent).Folders,
                f => string.Equals(f.Name, name, StringComparison.Ordinal));
        }
        finally
        {
            if (created != null)
            {
                var swept = commands.Delete(created, confirm: true);
                output.WriteLine($"Sweep: success={swept.Success} {swept.ErrorMessage}");
                Assert.True(swept.Success, swept.ErrorMessage);
            }
        }

        Assert.DoesNotContain(
            commands.ListChildren(parent).Folders,
            f => string.Equals(f.Name, name, StringComparison.Ordinal));
    }

    private static string? FindInDeletedItems(MailCommands commands, string subject)
    {
        var listing = commands.List(folder: "deleted", maxCount: 50, subjectContains: subject);
        Assert.True(listing.Success, listing.ErrorMessage);

        return listing.Messages
            .FirstOrDefault(i => string.Equals(i.Subject, subject, StringComparison.Ordinal))
            ?.EntryId;
    }

    private static string ResolveScratchParent(FolderCommands commands)
    {
        var probe = commands.ListDefault();
        Skip.If(!probe.Success, probe.ErrorMessage ?? "Outlook is not available.");

        var inbox = commands.ResolvePath("inbox");
        Skip.If(!inbox.Success, inbox.ErrorMessage);
        Skip.If(string.IsNullOrWhiteSpace(inbox.FolderPath), "Inbox reported no usable folder path.");

        return inbox.FolderPath!;
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
