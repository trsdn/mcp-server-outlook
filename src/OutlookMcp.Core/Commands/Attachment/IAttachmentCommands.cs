using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Attachment;

[ServiceCategory("attachment")]
[McpTool("attachment", Title = "Outlook Attachment Operations", Destructive = true, Category = "mail",
    Description = "Inspect and save attachments from a selected Outlook mail item without opening a persistent session. "
    + "Use list to inspect attachments on the active mail item or a specific mail entry id. "
    + "Use save to export one or all attachments to disk with explicit overwrite control. "
    + "Use add and remove to mutate attachments on Outlook draft items by entry id or active draft context — these actions modify the mail item and cannot be undone. "
    + "remove requires confirm=true and is refused without it: an attachment has no Deleted Items to be recovered from, "
    + "so removing it destroys the only copy the message holds.")]
public interface IAttachmentCommands
{
    [ServiceAction("list", Destructive = false)]
    AttachmentListResult List(
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("save", Destructive = false)]
    AttachmentSaveResult Save(
        string destinationDirectory,
        int attachmentIndex = 0,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool overwrite = false);

    /// <summary>
    /// Adds a file to a draft as an attachment.
    /// </summary>
    /// <param name="useActiveMail">Off by default. A mutating action must not fall back to whatever the user has selected in Outlook: the caller chooses the verb and the selection would silently choose the object. Pass an explicit <c>mailEntryId</c>, or set this to true when "the draft I am looking at" is genuinely what you mean.</param>
    [ServiceAction("add", Destructive = true)]
    AttachmentMutationResult Add(
        string filePath,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = false);

    /// <summary>
    /// Removes an attachment from a draft.
    ///
    /// <para>
    /// <b>Requires <paramref name="confirm"/>.</b> An attachment has no Deleted Items of its own, so
    /// this destroys the only copy the message holds. Call <c>list</c> first and confirm the index
    /// is the one the user meant.
    /// </para>
    /// </summary>
    /// <param name="useActiveMail">Off by default, for the same reason as <c>add</c>.</param>
    /// <param name="confirm">Must be true. Without it the call is refused and nothing is touched.</param>
    [ServiceAction("remove", Destructive = true)]
    AttachmentMutationResult Remove(
        int attachmentIndex,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = false,
        bool confirm = false);
}
