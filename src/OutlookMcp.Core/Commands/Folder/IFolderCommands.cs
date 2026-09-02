using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Folder;

[ServiceCategory("folder")]
[McpTool("folder", Title = "Outlook Folder Operations", Destructive = false, Category = "folder",
    Description = "Inspect Outlook mailbox folders without opening a persistent session. "
    + "Use list-default to enumerate important default folders such as Inbox, Drafts, Sent Items, Calendar, and Contacts. "
    + "Use list-children to enumerate child folders from the current folder, a default Outlook folder role, or an explicit Outlook folder path. "
    + "Use resolve-path to normalize a folder identifier, and list-items to inspect mixed Outlook items inside a resolved folder.")]
public interface IFolderCommands
{
    [ServiceAction("list-default")]
    OutlookFolderListResult ListDefault(bool includeItemCounts = false);

    [ServiceAction("list-children")]
    OutlookFolderListResult ListChildren(
        string? parentFolder = null,
        bool includeItemCounts = false);

    [ServiceAction("resolve-path")]
    OutlookFolderResolveResult ResolvePath(
        string? folder = null,
        bool includeItemCount = true);

    [ServiceAction("list-items")]
    OutlookFolderItemListResult ListItems(
        string? folder = null,
        int maxCount = 25,
        bool includePreview = false);
}
