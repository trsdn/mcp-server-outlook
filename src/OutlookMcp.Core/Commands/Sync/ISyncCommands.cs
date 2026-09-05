using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Sync;

[ServiceCategory("sync")]
[McpTool("sync", Title = "Outlook Send/Receive (Sync) Operations", Destructive = true, Category = "settings",
    Description = "Force a Send/Receive on the running Outlook, and inspect the Send/Receive groups. "
    + "Use list-groups to see the profile's Send/Receive groups and the Exchange connection mode "
    + "(cached vs online). Use send-receive to trigger a synchronisation before a critical read, or "
    + "to flush a queued outgoing message from the Outbox. "
    + "IMPORTANT: send-receive is ASYNCHRONOUS. Outlook's SyncObject.Start() returns immediately and "
    + "the synchronisation completes later on a background thread; this action reports that a sync was "
    + "STARTED, never that the mailbox is now current. Do not assume freshly-synced items are readable "
    + "the instant this returns. In pure Online (non-cached) mode there is usually nothing to "
    + "synchronise, and both actions report that distinction rather than silently doing nothing.")]
public interface ISyncCommands
{
    /// <summary>
    /// Lists the Send/Receive groups defined in the current Outlook profile (from
    /// <c>Namespace.SyncObjects</c>) and reports the default Exchange account's connection mode.
    /// Read-only. An empty list is a legitimate state — in pure Online mode there is nothing to
    /// synchronise — and is reported via <c>count</c> and <c>cacheMode</c> rather than treated as an
    /// error.
    /// </summary>
    [ServiceAction("list-groups")]
    OutlookSyncGroupListResult ListGroups();

    /// <summary>
    /// Starts a Send/Receive so the caller can make the mailbox current before a critical read or
    /// flush a queued outgoing message. When <paramref name="groupName"/> is given, only that group
    /// is started; otherwise every Send/Receive group is started.
    ///
    /// <para>
    /// ASYNCHRONOUS: this returns as soon as the sync has been <i>started</i>. Outlook performs the
    /// actual synchronisation on a background thread and signals completion through COM events that
    /// this out-of-process server does not block the shared Outlook thread to await. The result's
    /// <c>started</c> flag says a sync began; it does NOT promise the mailbox is up to date on
    /// return. If no Send/Receive groups exist (e.g. Online mode), <c>started</c> is false, the
    /// operation still succeeds, and <c>note</c> explains why.
    /// </para>
    /// </summary>
    /// <param name="groupName">Optional Send/Receive group name (as shown by list-groups, e.g. "All Accounts"). Omit to start all groups. Matching is case-insensitive; an unknown name is an error that lists the available groups.</param>
    [ServiceAction("send-receive", Destructive = true)]
    OutlookSyncStartResult SendReceive(string? groupName = null);
}
