using OutlookMcp.Core.Commands.Attachment;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.Folder;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// The confirmation gates that refuse an irreversible action before any Outlook COM call happens
/// (#9).
///
/// <para>
/// Only the gates that are genuinely pure live here. Each one is a guard clause evaluated on the
/// caller's arguments alone and returning before <c>OutlookInteropRunner.Execute</c> is reached, so
/// it satisfies every condition of the ADR-001 exception: no COM object is touched, real or
/// substituted, and the test fails if the guard is removed rather than only if .NET is broken.
/// </para>
///
/// <para>
/// The gates that <i>do</i> need Outlook - refusing a second delete of an item already sitting in
/// Deleted Items, which cannot be known without reading the item's parent folder - are covered by
/// <c>OutlookDeleteConfirmationTests</c> as integration tests, because that is the only way to
/// cover them honestly.
/// </para>
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "ConfirmationGate")]
[Trait("Speed", "Fast")]
[Trait("RequiresOutlook", "false")]
public class DestructiveConfirmationGateTests
{
    /// <summary>
    /// Deleting a folder takes every message and every subfolder in it, and in a store with no
    /// Deleted Items folder it is not a recycle-bin operation at all. Gated.
    /// </summary>
    [Fact]
    public void FolderDelete_WithoutConfirm_IsRefusedAndNeverTouchesOutlook()
    {
        var commands = new FolderCommands();

        var result = commands.Delete("Inbox/some-folder-that-need-not-exist", confirm: false);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("confirm=true", result.ErrorMessage!, StringComparison.Ordinal);

        // Refused *by the gate*, not by folder resolution: proof the guard runs before any COM call.
        Assert.DoesNotContain("could not be resolved", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The gate must not swallow the pre-existing argument checks. A blank folder was already
    /// refused for its own reason and must still be refused for that reason, not for the new one.
    /// </summary>
    [Fact]
    public void FolderDelete_WithBlankFolder_IsStillRefusedForTheOriginalReason()
    {
        var commands = new FolderCommands();

        var result = commands.Delete("   ", confirm: true);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("A folder is required", result.ErrorMessage!, StringComparison.Ordinal);
    }

    /// <summary>
    /// An attachment has no Deleted Items to be recovered from. Removing one destroys the only copy
    /// the message holds, so it is gated even though the message itself survives.
    /// </summary>
    [Fact]
    public void AttachmentRemove_WithoutConfirm_IsRefusedAndNeverTouchesOutlook()
    {
        var commands = new AttachmentCommands();

        var result = commands.Remove(attachmentIndex: 1, mailEntryId: "some-entry-id", confirm: false);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("confirm=true", result.ErrorMessage!, StringComparison.Ordinal);
    }

    /// <summary>
    /// The index check came first before the gate existed and must still come first: an index of 0
    /// is a caller mistake worth naming, and answering it with "pass confirm=true" would send the
    /// caller to fix the wrong thing.
    /// </summary>
    [Fact]
    public void AttachmentRemove_WithAnInvalidIndex_IsStillRefusedForTheOriginalReason()
    {
        var commands = new AttachmentCommands();

        var result = commands.Remove(attachmentIndex: 0, mailEntryId: "some-entry-id", confirm: true);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("attachmentIndex", result.ErrorMessage!, StringComparison.Ordinal);
    }

    /// <summary>
    /// Cancelling one occurrence of a recurring series writes a deletion exception into the
    /// recurrence pattern. Nothing lands in Deleted Items, so there is nothing to restore - unlike
    /// deleting the series itself, which is an ordinary recoverable soft delete.
    /// </summary>
    [Fact]
    public void CalendarDeleteOccurrence_WithoutConfirm_IsRefusedAndNeverTouchesOutlook()
    {
        var commands = new CalendarCommands();

        var result = commands.DeleteAppointment(
            entryId: "some-entry-id",
            occurrenceDate: "2026-03-07",
            confirm: false);

        Assert.False(result.Success);
        Assert.False(result.Deleted);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("confirm=true", result.ErrorMessage!, StringComparison.Ordinal);
        Assert.Contains("occurrence", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }
}
