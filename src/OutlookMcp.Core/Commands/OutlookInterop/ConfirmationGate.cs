using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.OutlookInterop;

/// <summary>
/// The wording of the confirmation gates, in one place.
///
/// <para>
/// <b>What is gated and what is not, and why (#9).</b> A confirmation gate is only worth having
/// where the action cannot be undone. Applying one to a recoverable action trains the caller to
/// pass <c>confirm=true</c> reflexively, which is exactly how the gate on the irreversible action
/// stops being read. So the line is drawn at recoverability, and drawn explicitly:
/// </para>
///
/// <list type="table">
///   <listheader><term>Action</term><description>What Outlook actually does</description></listheader>
///   <item>
///     <term><c>mail.delete</c>, <c>contact.delete</c>, <c>task.delete</c>,
///     <c>calendar.delete-appointment</c> (whole series)</term>
///     <description><b>Not gated.</b> <c>Item.Delete</c> moves the item to the store's Deleted Items
///     folder. The user can restore it from the Outlook UI, so the agent asking first adds
///     ceremony, not safety.</description>
///   </item>
///   <item>
///     <term>the same four, when the item is <i>already</i> in Deleted Items</term>
///     <description><b>Gated.</b> There is no second recycle bin. This delete destroys the item, and
///     the caller cannot tell the two cases apart from the entry id alone - see
///     <see cref="OutlookInteropRunner.IsInDeletedItems"/>.</description>
///   </item>
///   <item>
///     <term><c>mail.move</c></term>
///     <description><b>Not gated.</b> Reversible by moving the item back. The entry id changes, so
///     the response reports the new one.</description>
///   </item>
///   <item>
///     <term><c>folder.delete</c></term>
///     <description><b>Gated.</b> Every message and every subfolder goes with the folder, and in a
///     store with no Deleted Items folder it is not a recycle-bin operation at all.</description>
///   </item>
///   <item>
///     <term><c>calendar.delete-appointment</c> with <c>occurrenceDate</c></term>
///     <description><b>Gated.</b> Cancelling one occurrence writes a deletion exception into the
///     recurrence pattern. Nothing lands in Deleted Items, so there is nothing to restore.</description>
///   </item>
///   <item>
///     <term><c>attachment.remove</c></term>
///     <description><b>Gated.</b> An attachment has no Deleted Items of its own; removing it
///     destroys the only copy the message holds.</description>
///   </item>
///   <item>
///     <term><c>mail.send</c></term>
///     <description><b>Gated</b> (pre-dates this; see #29). A sent message cannot be recalled and
///     the effect is outside the mailbox entirely.</description>
///   </item>
/// </list>
/// </summary>
internal static class ConfirmationGate
{
    /// <summary>
    /// The refusal for an item delete that would be permanent because the item is already in
    /// Deleted Items.
    /// </summary>
    /// <param name="itemDescription">What is about to go, in the caller's terms - "Outlook mail item", "contact".</param>
    /// <param name="action">The action name as the caller invoked it, e.g. <c>mail delete</c>.</param>
    internal static string AlreadyInDeletedItems(string itemDescription, string action) =>
        $"Deleting this {itemDescription} requires confirm=true. It is already in Deleted Items, so "
        + "this is a permanent delete rather than the usual move to Deleted Items, and nothing will "
        + $"be left to restore (#9). Call {action} again with confirm=true once you have confirmed "
        + "with the user that this is the item they meant to destroy.";

    internal static string FolderDelete(string folder) =>
        $"Deleting the folder '{folder}' requires confirm=true. This is a deliberate confirmation "
        + "gate for an irreversible action (#9): every message and every subfolder inside it goes "
        + "with it, and in a store without a Deleted Items folder it is gone outright rather than "
        + "recoverable. Call folder delete again with confirm=true once you have listed its children "
        + "and told the user what will be lost.";

    internal static string DeleteOccurrence() =>
        "Cancelling one occurrence of a recurring series requires confirm=true. This is a deliberate "
        + "confirmation gate for an irreversible action (#9): the cancelled occurrence is recorded as "
        + "an exception in the recurrence pattern rather than moved to Deleted Items, so nothing is "
        + "left to restore. Call delete-appointment again with confirm=true, or omit occurrenceDate "
        + "to delete the whole series, which does go to Deleted Items and can be recovered.";

    internal static string AttachmentRemove() =>
        "Removing an attachment requires confirm=true. This is a deliberate confirmation gate for an "
        + "irreversible action (#9): an attachment has no Deleted Items to be recovered from, so this "
        + "destroys the only copy the message holds. Call attachment remove again with confirm=true "
        + "once you have listed the attachments and confirmed the index is the one the user meant.";

    /// <summary>
    /// Whether an item about to be deleted needs confirmation, i.e. whether the delete would be
    /// permanent rather than a move to Deleted Items.
    /// </summary>
    internal static bool RequiresConfirmationToDelete(bool confirm, Outlook.MAPIFolder? parentFolder) =>
        !confirm && OutlookInteropRunner.IsInDeletedItems(parentFolder);

    /// <summary>
    /// The refusal for emptying a folder without confirmation. Emptying clears every item in the
    /// folder in one call, so unlike deleting a single message it is gated even though the items go
    /// to Deleted Items (#15).
    /// </summary>
    /// <param name="folder">The folder as the caller named it.</param>
    internal static string FolderEmpty(string folder) =>
        $"Emptying the folder '{folder}' requires confirm=true. This is a deliberate confirmation "
        + "gate for a bulk destructive action (#15): every item in the folder is moved to Deleted "
        + "Items in one step, which is not the same as deleting a single message. Subfolders are left "
        + "untouched. Call folder empty again with confirm=true once you have listed the folder's "
        + "items and told the user how many will be cleared.";
}
