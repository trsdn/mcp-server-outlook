using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Sync;

/// <summary>
/// Send/Receive (synchronisation) operations backed by <c>Namespace.SyncObjects</c>.
/// </summary>
public class SyncCommands : ISyncCommands
{
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookSyncGroupListResult ListGroups()
    {
        return OutlookInteropRunner.Execute(
            "OutlookSyncListGroups",
            (application, session) =>
            {
                Outlook.SyncObjects? syncObjects = null;
                try
                {
                    string connectionMode = GetConnectionMode(session, out string cacheMode);

                    syncObjects = session.SyncObjects;
                    int count = syncObjects?.Count ?? 0;

                    var groups = new List<OutlookSyncGroupInfo>(count);
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.SyncObject? syncObject = null;
                        try
                        {
                            syncObject = syncObjects![i];
                            groups.Add(new OutlookSyncGroupInfo { Name = syncObject.Name ?? string.Empty });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref syncObject);
                        }
                    }

                    return new OutlookSyncGroupListResult
                    {
                        Success = true,
                        Groups = groups,
                        Count = groups.Count,
                        ExchangeConnectionMode = connectionMode,
                        CacheMode = cacheMode
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref syncObjects);
                }
            },
            ex => new OutlookSyncGroupListResult
            {
                Success = false,
                ErrorMessage = $"Failed to list Outlook Send/Receive groups: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookSyncStartResult SendReceive(string? groupName = null)
    {
        bool targetSpecificGroup = !string.IsNullOrWhiteSpace(groupName);

        // Declared out here so the onException delegate can still report which groups were already
        // started if a later group throws mid-loop. Start() is asynchronous and fire-and-forget, so
        // by the time an exception surfaces the earlier groups are genuinely running; a failure result
        // that claimed nothing started would deny side effects that have happened and could invite a
        // retry that double-flushes the Outbox. (Rule 1b: we do NOT catch inside the action lambda to
        // build an error result — we let the exception reach onException, which reads this state.)
        var started = new List<string>();
        string connectionMode = "unknown";

        return OutlookInteropRunner.Execute(
            "OutlookSyncSendReceive",
            (application, session) =>
            {
                Outlook.SyncObjects? syncObjects = null;
                try
                {
                    connectionMode = GetConnectionMode(session, out _);

                    syncObjects = session.SyncObjects;
                    int count = syncObjects?.Count ?? 0;

                    // Collect the names first so an unknown-group error can list what is available
                    // without starting anything.
                    var available = new List<string>(count);
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.SyncObject? probe = null;
                        try
                        {
                            probe = syncObjects![i];
                            available.Add(probe.Name ?? string.Empty);
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref probe);
                        }
                    }

                    // Validate a requested group BEFORE the empty-collection shortcut. A named group
                    // that does not exist is an error regardless of how many groups the profile has —
                    // including zero. Doing this first stops "send-receive <unknown>" from spuriously
                    // succeeding on a profile that exposes no Send/Receive groups.
                    if (targetSpecificGroup &&
                        !available.Any(n => string.Equals(n, groupName, StringComparison.OrdinalIgnoreCase)))
                    {
                        return new OutlookSyncStartResult
                        {
                            Success = false,
                            Started = false,
                            ExchangeConnectionMode = connectionMode,
                            ErrorMessage = available.Count == 0
                                ? $"No Send/Receive group named '{groupName}' was found; this profile has "
                                    + "no Send/Receive groups (this is normal for a pure Online, non-cached "
                                    + "Exchange connection)."
                                : $"No Send/Receive group named '{groupName}' was found. "
                                    + $"Available groups: {string.Join(", ", available)}."
                        };
                    }

                    if (count == 0)
                    {
                        // No group was requested and none exist. Nothing to synchronise. Not an error:
                        // report it plainly so an Online-mode caller is not misled into thinking a sync
                        // happened.
                        return new OutlookSyncStartResult
                        {
                            Success = true,
                            Started = false,
                            ExchangeConnectionMode = connectionMode,
                            Note = "No Send/Receive groups are defined in this profile, so there is "
                                + "nothing to synchronise. This is normal for a pure Online (non-cached) "
                                + "Exchange connection."
                        };
                    }

                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.SyncObject? syncObject = null;
                        try
                        {
                            syncObject = syncObjects![i];
                            string name = syncObject.Name ?? string.Empty;

                            if (targetSpecificGroup &&
                                !string.Equals(name, groupName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            // Start() is asynchronous: it returns immediately and Outlook performs the
                            // synchronisation on a background thread. We deliberately do NOT wait on the
                            // SyncStart/Progress/SyncEnd events, because the only thread we could wait on
                            // is the single process-wide dispatcher STA thread, and blocking it would
                            // wedge every other Outlook operation in the process.
                            syncObject.Start();
                            started.Add(name);
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref syncObject);
                        }
                    }

                    return new OutlookSyncStartResult
                    {
                        Success = true,
                        Started = started.Count > 0,
                        StartedGroups = started,
                        ExchangeConnectionMode = connectionMode,
                        Note = "Send/Receive was started asynchronously and completes in the background. "
                            + "The mailbox is not guaranteed to be current the instant this returns."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref syncObjects);
                }
            },
            ex => new OutlookSyncStartResult
            {
                Success = false,
                // Report the groups that were already started before the failure. They are running
                // asynchronously and cannot be un-started; hiding them would invite a double-flush retry.
                Started = started.Count > 0,
                StartedGroups = started,
                ExchangeConnectionMode = connectionMode,
                ErrorMessage = started.Count > 0
                    ? $"Send/Receive partially started ({string.Join(", ", started)}) before failing: {ex.Message}"
                    : $"Failed to start Outlook Send/Receive: {ex.Message}"
            });
    }

    /// <summary>
    /// Reads <c>Namespace.ExchangeConnectionMode</c> and derives a coarse cache-mode label. A profile
    /// with no Exchange account (POP/IMAP only, or none) reports <c>olNoExchange</c>; the property can
    /// also throw on some profiles, which is treated as "unknown" rather than failing the whole call.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string GetConnectionMode(Outlook.NameSpace session, out string cacheMode)
    {
        Outlook.OlExchangeConnectionMode mode;
        try
        {
            mode = session.ExchangeConnectionMode;
        }
        catch
        {
            // Reading the connection mode is a best-effort annotation, not the operation itself.
            // Its absence must not fail a Send/Receive or a group listing.
            cacheMode = "unknown";
            return "unknown";
        }

        cacheMode = mode switch
        {
            Outlook.OlExchangeConnectionMode.olNoExchange => "none",
            Outlook.OlExchangeConnectionMode.olOnline => "online",
            Outlook.OlExchangeConnectionMode.olOffline => "offline",
            Outlook.OlExchangeConnectionMode.olDisconnected => "disconnected",
            Outlook.OlExchangeConnectionMode.olCachedDisconnected => "disconnected",
            Outlook.OlExchangeConnectionMode.olCachedOffline => "offline",
            Outlook.OlExchangeConnectionMode.olCachedConnectedHeaders => "cached",
            Outlook.OlExchangeConnectionMode.olCachedConnectedDrizzle => "cached",
            Outlook.OlExchangeConnectionMode.olCachedConnectedFull => "cached",
            _ => "unknown"
        };

        return mode.ToString();
    }
}
