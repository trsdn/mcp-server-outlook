using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Attachment;

[ServiceCategory("attachment")]
[McpTool("attachment", Title = "Outlook Attachment Operations", Destructive = true, Category = "mail",
    Description = "Inspect and save attachments from a selected Outlook mail item without opening a persistent session. "
    + "Use list to inspect attachments on the active mail item or a specific mail entry id. "
    + "Use save to export one or all attachments to disk with explicit overwrite control. "
    + "Use add and remove to mutate attachments on Outlook draft items by entry id or active draft context — these actions modify the mail item and cannot be undone.")]
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

    [ServiceAction("add", Destructive = true)]
    AttachmentMutationResult Add(
        string filePath,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("remove", Destructive = true)]
    AttachmentMutationResult Remove(
        int attachmentIndex,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true);
}
