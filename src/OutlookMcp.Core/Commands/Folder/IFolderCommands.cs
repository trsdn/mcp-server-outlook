using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Folder;

[ServiceCategory("folder")]
[McpTool("folder", Title = "Outlook Folder Operations", Destructive = false, Category = "folder",
    Description = "Inspect Outlook mailbox folders without opening a persistent session. "
    + "Use list-default to enumerate important default folders such as Inbox, Drafts, Sent Items, Calendar, and Contacts. "
    + "Use list-children to enumerate child folders from the current folder, a default Outlook folder role, or an explicit Outlook folder path. "
    + "Use resolve-path to normalize a folder identifier, and list-items to inspect mixed Outlook items inside a resolved folder. "
    + "Use list-stores to discover every mailbox, archive and data file in the profile. A profile often holds more than one, "
    + "and every one of them has its own Inbox, so an unqualified request only ever reaches the default delivery store. "
    + "Pass a storeId from list-stores to list-default to read a specific mailbox; folder results name the store they came from. "
    + "Use open-shared to reach another person's Inbox or Calendar when they have granted access, without adding their mailbox to the profile; "
    + "it returns a folder path usable with the mail and calendar tools. "
    + "Use create, rename, move and delete to change the folder tree - create takes a parent folder and a name, "
    + "and delete removes the folder together with everything filed in it. "
    + "Default folders such as Inbox, Sent Items and Calendar, and store roots, are refused for rename, move and delete: "
    + "Outlook itself allows those and they are not recoverable.")]
public interface IFolderCommands
{
    [ServiceAction("list-default")]
    OutlookFolderListResult ListDefault(bool includeItemCounts = false, string? storeId = null);

    [ServiceAction("list-stores")]
    OutlookStoreListResult ListStores();

    [ServiceAction("open-shared")]
    OutlookFolderResolveResult OpenShared(string? address = null, string? role = null);

    [ServiceAction("create", Destructive = true)]
    OutlookFolderResolveResult Create(string? parentFolder = null, string? name = null);

    [ServiceAction("rename", Destructive = true)]
    OutlookFolderResolveResult Rename(string? folder = null, string? name = null);

    [ServiceAction("move", Destructive = true)]
    OutlookFolderResolveResult Move(string? folder = null, string? destinationFolder = null);

    [ServiceAction("delete", Destructive = true)]
    OutlookFolderResolveResult Delete(string? folder = null);

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
