using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Folder;

[ServiceCategory("folder")]
[NoSession]
[McpTool("folder", Title = "Outlook Folder Operations", Destructive = false, Category = "folder",
    Description = "Inspect Outlook mailbox folders without opening a persistent session. "
    + "Use list-default to enumerate important default folders such as Inbox, Drafts, Sent Items, Calendar, and Contacts. "
    + "Use list-children to enumerate child folders from the current folder, a default Outlook folder role, or an explicit Outlook folder path.")]
public interface IFolderCommands
{
    [ServiceAction("list-default")]
    OutlookFolderListResult ListDefault(bool includeItemCounts = false);

    [ServiceAction("list-children")]
    OutlookFolderListResult ListChildren(
        string? parentFolder = null,
        bool includeItemCounts = false);
}
