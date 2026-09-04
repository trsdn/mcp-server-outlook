namespace OutlookMcp.Core.Commands.OutlookInterop;

/// <summary>
/// The wording of the confirmation gates, in one place.
///
/// <para>
/// A confirmation gate is only worth having where the action cannot be walked back cheaply.
/// Applying one to a recoverable action trains the caller to pass <c>confirm=true</c> reflexively,
/// which is exactly how the gate on the irreversible action stops being read. So the line is drawn
/// at how much goes and how hard it is to get back, and drawn explicitly per action.
/// </para>
///
/// <para>
/// <b>Note (#15 / #9).</b> The confirm=true pattern was introduced by the change tracked in #9,
/// which adds this same class with a broader set of gate messages. This copy carries only the
/// folder-empty message so the folder-empty slice is self-contained; when the two land together the
/// class bodies are merged into one.
/// </para>
/// </summary>
internal static class ConfirmationGate
{
    /// <summary>
    /// The refusal for emptying a folder without confirmation. Emptying clears every item in the
    /// folder in one call, so unlike deleting a single message it is gated even though the items go
    /// to Deleted Items.
    /// </summary>
    /// <param name="folder">The folder as the caller named it.</param>
    internal static string FolderEmpty(string folder) =>
        $"Emptying the folder '{folder}' requires confirm=true. This is a deliberate confirmation "
        + "gate for a bulk destructive action (#15): every item in the folder is moved to Deleted "
        + "Items in one step, which is not the same as deleting a single message. Subfolders are left "
        + "untouched. Call folder empty again with confirm=true once you have listed the folder's "
        + "items and told the user how many will be cleared.";
}
