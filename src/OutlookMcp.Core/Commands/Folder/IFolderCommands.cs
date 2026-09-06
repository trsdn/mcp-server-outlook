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
    + "and delete removes the folder together with everything filed in it. delete requires confirm=true and is "
    + "refused without it, because nothing about a deleted folder is recoverable in every store. "
    + "Use empty to clear a folder's own items (moving them to Deleted Items) while keeping the folder and its subfolders; "
    + "empty requires confirm=true. "
    + "Default folders such as Inbox, Sent Items and Calendar, and store roots, are refused for rename, move, delete and empty: "
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

    /// <summary>
    /// Deletes a folder and everything filed in it.
    ///
    /// <para>
    /// <b>Requires <paramref name="confirm"/>.</b> This is the one folder operation with no way
    /// back: every message and every subfolder goes with the folder, and in a store without a
    /// Deleted Items folder it is gone outright rather than moved there. List the folder's children
    /// and tell the user what will be lost before passing <c>confirm=true</c>.
    /// </para>
    /// </summary>
    /// <param name="folder">The folder to delete. There is no default: falling back to the current folder would delete whatever the user happens to have selected.</param>
    /// <param name="confirm">Must be true. Without it the call is refused and nothing is touched.</param>
    [ServiceAction("delete", Destructive = true)]
    OutlookFolderResolveResult Delete(string? folder = null, bool confirm = false);

    /// <summary>
    /// Deletes the items inside a folder while keeping the folder itself.
    ///
    /// <para>
    /// <b>Semantics, stated so an agent never has to guess.</b> Empty removes the folder's own items
    /// only - it moves each one to the store's Deleted Items folder (the same as deleting a single
    /// mail item), so in a normal mailbox the contents can still be recovered from Deleted Items.
    /// <b>Subfolders and everything inside them are left untouched:</b> "empty the archive" clears the
    /// archive's own messages but keeps its sub-folders and their contents. To remove a subfolder too,
    /// delete it explicitly.
    /// </para>
    ///
    /// <para>
    /// <b>This is the most dangerous action here.</b> Emptying the Inbox in one call is not the same
    /// as deleting one message, so default and special folders (Inbox, Sent Items, Drafts, Deleted
    /// Items, Calendar, Contacts, Tasks, Notes, Junk) and store roots are refused outright, across
    /// every store. It also requires <c>confirm=true</c>: without it the folder is resolved, the guard
    /// is checked, and then the call is refused so the agent has to confirm with the user first. The
    /// result reports how many items were removed, so an empty folder (0 removed, success) is
    /// distinguishable from a refusal.
    /// </para>
    /// </summary>
    /// <param name="folder">The folder to empty, as a path or a resolved identifier. There is no default target. Default/special folders and store roots are refused.</param>
    /// <param name="confirm">Must be <c>true</c> to proceed. Left <c>false</c>, the call is refused after the folder is resolved and the guards are checked, so the agent can tell the user exactly what would be cleared before repeating the call.</param>
    [ServiceAction("empty", Destructive = true)]
    OutlookFolderResolveResult Empty(string? folder = null, bool confirm = false);

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
