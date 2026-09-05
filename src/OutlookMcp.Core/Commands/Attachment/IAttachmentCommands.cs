using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Attachment;

[ServiceCategory("attachment")]
[McpTool("attachment", Title = "Outlook Attachment Operations", Destructive = true, Category = "mail",
    Description = "Inspect and save attachments from a selected Outlook mail item without opening a persistent session. "
    + "Use list to inspect attachments on the active mail item or a specific mail entry id; each attachment reports a 1-based index and a fileName. "
    + "Use save to export attachments to disk: pick one by attachmentName (preferred - names are what list returned and are stable), or by attachmentIndex (1-based, matching list), or set allAttachments to export every one. "
    + "Attachment indices are 1-based, matching Outlook's collection: the first attachment is 1, not 0. attachmentIndex=0 means 'no index supplied' and is rejected - it is NOT a shortcut for 'all'; use allAttachments=true for that. "
    + "Use add and remove to mutate attachments on Outlook draft items by entry id or active draft context - these actions modify the mail item and cannot be undone. "
    + "remove requires confirm=true and is refused without it: an attachment has no Deleted Items to be recovered from, "
    + "so removing it destroys the only copy the message holds.")]
public interface IAttachmentCommands
{
    [ServiceAction("list", Destructive = false)]
    AttachmentListResult List(
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    /// <summary>
    /// Save one or more attachments from a mail item to a directory on disk.
    /// </summary>
    /// <param name="destinationDirectory">Directory to write the attachment file(s) into. Created if it does not exist.</param>
    /// <param name="attachmentIndex">1-based position of the single attachment to save, matching the 'index' field from attachment list. Outlook attachment collections are 1-based, so the first attachment is 1. The default 0 means 'no index supplied' — it is not a valid attachment and is NOT a shortcut for 'all'; set allAttachments=true to export every attachment instead.</param>
    /// <param name="attachmentName">File name of the single attachment to save, matching the 'fileName' field from attachment list. Preferred over attachmentIndex for automated callers because names are stable while indices shift as attachments are added or removed. If more than one attachment shares this name, all matching attachments are saved.</param>
    /// <param name="allAttachments">Save every attachment on the item. This is the explicit way to export all attachments; it replaces the former overloaded attachmentIndex=0. Provide exactly one of attachmentIndex, attachmentName or allAttachments.</param>
    /// <param name="overwrite">Overwrite existing files in the destination directory. When false, the save fails if a target file already exists.</param>
    [ServiceAction("save", Destructive = false)]
    AttachmentSaveResult Save(
        string destinationDirectory,
        int attachmentIndex = 0,
        string? attachmentName = null,
        bool allAttachments = false,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool overwrite = false);

    [ServiceAction("add", Destructive = true)]
    AttachmentMutationResult Add(
        string filePath,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    /// <summary>
    /// Removes an attachment from a draft.
    ///
    /// <para>
    /// <b>Requires <paramref name="confirm"/>.</b> An attachment has no Deleted Items of its own, so
    /// this destroys the only copy the message holds. Call <c>list</c> first and confirm the index
    /// is the one the user meant.
    /// </para>
    /// </summary>
    /// <param name="confirm">Must be true. Without it the call is refused and nothing is touched.</param>
    [ServiceAction("remove", Destructive = true)]
    AttachmentMutationResult Remove(
        int attachmentIndex,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool confirm = false);
}
