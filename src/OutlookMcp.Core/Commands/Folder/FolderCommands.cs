using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Folder;

public class FolderCommands : IFolderCommands
{
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> DefaultFolderAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["inbox"] = Outlook.OlDefaultFolders.olFolderInbox,
            ["drafts"] = Outlook.OlDefaultFolders.olFolderDrafts,
            ["sent"] = Outlook.OlDefaultFolders.olFolderSentMail,
            ["outbox"] = Outlook.OlDefaultFolders.olFolderOutbox,
            ["deleted"] = Outlook.OlDefaultFolders.olFolderDeletedItems,
            ["calendar"] = Outlook.OlDefaultFolders.olFolderCalendar,
            ["contacts"] = Outlook.OlDefaultFolders.olFolderContacts,
            ["tasks"] = Outlook.OlDefaultFolders.olFolderTasks,
            ["notes"] = Outlook.OlDefaultFolders.olFolderNotes,
            ["junk"] = Outlook.OlDefaultFolders.olFolderJunk
        };

    private static readonly (string Role, Outlook.OlDefaultFolders Folder)[] DefaultFolders =
    [
        ("inbox", Outlook.OlDefaultFolders.olFolderInbox),
        ("drafts", Outlook.OlDefaultFolders.olFolderDrafts),
        ("sent", Outlook.OlDefaultFolders.olFolderSentMail),
        ("outbox", Outlook.OlDefaultFolders.olFolderOutbox),
        ("deleted", Outlook.OlDefaultFolders.olFolderDeletedItems),
        ("calendar", Outlook.OlDefaultFolders.olFolderCalendar),
        ("contacts", Outlook.OlDefaultFolders.olFolderContacts),
        ("tasks", Outlook.OlDefaultFolders.olFolderTasks),
        ("notes", Outlook.OlDefaultFolders.olFolderNotes),
        ("junk", Outlook.OlDefaultFolders.olFolderJunk)
    ];

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderListResult ListDefault(bool includeItemCounts = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderListDefault",
            (application, session) =>
            {
                var result = new OutlookFolderListResult
                {
                    Success = true
                };

                foreach (var entry in DefaultFolders)
                {
                    Outlook.MAPIFolder? folder = null;
                    Outlook.Items? items = null;

                    try
                    {
                        folder = session.GetDefaultFolder(entry.Folder);
                        int? itemCount = null;
                        if (includeItemCounts)
                        {
                            items = folder.Items;
                            itemCount = items.Count;
                        }

                        result.Folders.Add(new OutlookFolderInfo
                        {
                            Role = entry.Role,
                            Available = true,
                            Name = folder.Name,
                            FolderPath = OutlookInteropRunner.GetFolderPath(folder),
                            ItemCount = itemCount
                        });
                    }
                    catch
                    {
                        result.Folders.Add(new OutlookFolderInfo
                        {
                            Role = entry.Role,
                            Available = false
                        });
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref items);
                        OutlookInteropRunner.ReleaseComObject(ref folder);
                    }
                }

                return result;
            },
            ex => new OutlookFolderListResult
            {
                Success = false,
                ErrorMessage = $"Failed to read Outlook default folders: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderListResult ListChildren(
        string? parentFolder = null,
        bool includeItemCounts = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderListChildren",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Folders? childFolders = null;

                try
                {
                    resolvedFolder = OutlookInteropRunner.ResolveFolder(
                        application,
                        session,
                        parentFolder,
                        DefaultFolderAliases,
                        ref explorer);

                    if (resolvedFolder == null)
                    {
                        return new OutlookFolderListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(parentFolder)
                        };
                    }

                    var result = new OutlookFolderListResult
                    {
                        Success = true
                    };

                    childFolders = resolvedFolder.Folders;
                    int childCount = childFolders.Count;
                    for (int index = 1; index <= childCount; index++)
                    {
                        Outlook.MAPIFolder? childFolder = null;
                        Outlook.Items? items = null;

                        try
                        {
                            childFolder = childFolders[index];
                            int? itemCount = null;
                            if (includeItemCounts)
                            {
                                items = childFolder.Items;
                                itemCount = items.Count;
                            }

                            result.Folders.Add(new OutlookFolderInfo
                            {
                                Role = SafeGet(() => childFolder.Name) ?? $"child-{index}",
                                Available = true,
                                Name = SafeGet(() => childFolder.Name),
                                FolderPath = OutlookInteropRunner.GetFolderPath(childFolder),
                                ItemCount = itemCount
                            });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref items);
                            OutlookInteropRunner.ReleaseComObject(ref childFolder);
                        }
                    }

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref childFolders);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookFolderListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook child folders: {ex.Message}"
            });
    }

    private static string? SafeGet(Func<string?> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static string BuildUnknownFolderMessage(string? folder)
    {
        const string supportedFolders = "current, inbox, drafts, sent, outbox, deleted, calendar, contacts, tasks, notes, junk, or an Outlook folder path";
        return string.IsNullOrWhiteSpace(folder)
            ? $"Could not resolve the Outlook folder. Supported folder values: {supportedFolders}."
            : $"Unsupported Outlook folder '{folder}'. Supported folder values: {supportedFolders}.";
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult ResolvePath(
        string? folder = null,
        bool includeItemCount = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderResolvePath",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Folders? childFolders = null;
                Outlook.Items? items = null;

                try
                {
                    resolvedFolder = OutlookInteropRunner.ResolveFolder(
                        application,
                        session,
                        folder,
                        DefaultFolderAliases,
                        ref explorer);

                    if (resolvedFolder == null)
                    {
                        return new OutlookFolderResolveResult
                        {
                            Success = false,
                            RequestedFolder = folder,
                            Resolved = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    childFolders = resolvedFolder.Folders;
                    int? itemCount = null;
                    if (includeItemCount)
                    {
                        items = resolvedFolder.Items;
                        itemCount = SafeGetInt(() => items.Count);
                    }

                    return new OutlookFolderResolveResult
                    {
                        Success = true,
                        RequestedFolder = folder,
                        Resolved = true,
                        Name = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        StoreId = SafeGet(() => resolvedFolder.StoreID),
                        DefaultRole = TryGetDefaultRole(folder),
                        ChildFolderCount = SafeGetInt(() => childFolders.Count),
                        ItemCount = itemCount
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref childFolders);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookFolderResolveResult
            {
                Success = false,
                RequestedFolder = folder,
                Resolved = false,
                ErrorMessage = $"Failed to resolve the Outlook folder: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderItemListResult ListItems(
        string? folder = null,
        int maxCount = 25,
        bool includePreview = false)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);

        return OutlookInteropRunner.Execute(
            "OutlookFolderListItems",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Items? items = null;

                try
                {
                    resolvedFolder = OutlookInteropRunner.ResolveFolder(
                        application,
                        session,
                        folder,
                        DefaultFolderAliases,
                        ref explorer);

                    if (resolvedFolder == null)
                    {
                        return new OutlookFolderItemListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = resolvedFolder.Items;
                    int totalItemCount = SafeGetInt(() => items.Count);

                    // Without this the cap below returns an arbitrary subset in store order, which
                    // a caller reads as "this is what is in the folder" (#91). Ordering is attempted
                    // newest-first and the property actually used is reported, because a folder of
                    // appointments or contacts has no received time and the honest answer there is a
                    // different ordering rather than a pretended one.
                    string? sortedBy = TrySortNewestFirst(items);

                    var result = new OutlookFolderItemListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        TotalItemCount = totalItemCount,
                        Truncated = totalItemCount > boundedMaxCount,
                        SortedBy = sortedBy,
                        SortDirection = sortedBy == null ? null : "descending"
                    };

                    for (int index = 1; index <= totalItemCount && result.Items.Count < boundedMaxCount; index++)
                    {
                        object? rawItem = null;

                        try
                        {
                            rawItem = items[index];
                            var info = CreateFolderItemInfo(rawItem, includePreview);
                            if (info != null)
                            {
                                result.Items.Add(info);
                            }
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ReturnedCount = result.Items.Count;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookFolderItemListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook folder items: {ex.Message}"
            });
    }

    /// <summary>
    /// Orders a folder's items newest-first, returning the property that worked.
    ///
    /// <para>
    /// <c>ReceivedTime</c> is preferred because it is what "newest" means for mail, but it does not
    /// exist on appointments, contacts or tasks and Outlook throws rather than ignoring it there.
    /// <c>LastModificationTime</c> exists on every item type and is the fallback. If both fail the
    /// caller is told the order is unknown rather than being handed store order dressed up as an
    /// ordering. See #91.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? TrySortNewestFirst(Outlook.Items items)
    {
        foreach ((string property, string reported) in SortCandidates)
        {
            try
            {
                items.Sort(property, true);
                return reported;
            }
            catch (COMException)
            {
                // Property not available on this folder's item types; try the next one.
            }
        }

        return null;
    }

    private static readonly (string Property, string Reported)[] SortCandidates =
    [
        ("[ReceivedTime]", "receivedTime"),
        ("[LastModificationTime]", "lastModificationTime")
    ];

    private static DateTimeOffset? SafeGetDateTimeOffset(Func<DateTime> getter)
    {
        try
        {
            DateTime value = getter();
            if (value == default)
            {
                return null;
            }

            return new DateTimeOffset(value);
        }
        catch
        {
            return null;
        }
    }

    private static int SafeGetInt(Func<int> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return 0;
        }
    }

    private static bool SafeGetBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return false;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static OutlookFolderItemInfo? CreateFolderItemInfo(object rawItem, bool includePreview)
    {
        if (rawItem is Outlook.MailItem mail)
        {
            return new OutlookFolderItemInfo
            {
                EntryId = SafeGet(() => mail.EntryID),
                StoreId = SafeGet(() => (mail.Parent as Outlook.MAPIFolder)?.StoreID),
                ItemType = "mail",
                MessageClass = SafeGet(() => mail.MessageClass),
                Subject = SafeGet(() => mail.Subject),
                Name = SafeGet(() => mail.SenderName),
                Preview = includePreview
                    ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body))
                    : null,
                ReceivedTime = SafeGetDateTimeOffset(() => mail.ReceivedTime),
                Unread = SafeGetBool(() => mail.UnRead)
            };
        }

        if (rawItem is Outlook.AppointmentItem appointment)
        {
            return new OutlookFolderItemInfo
            {
                EntryId = SafeGet(() => appointment.EntryID),
                StoreId = SafeGet(() => (appointment.Parent as Outlook.MAPIFolder)?.StoreID),
                ItemType = "appointment",
                MessageClass = SafeGet(() => appointment.MessageClass),
                Subject = SafeGet(() => appointment.Subject),
                Name = SafeGet(() => appointment.Organizer),
                Preview = includePreview
                    ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => appointment.Body))
                    : null,
                Start = SafeGetDateTimeOffset(() => appointment.Start),
                End = SafeGetDateTimeOffset(() => appointment.End)
            };
        }

        if (rawItem is Outlook.ContactItem contact)
        {
            return new OutlookFolderItemInfo
            {
                EntryId = SafeGet(() => contact.EntryID),
                StoreId = SafeGet(() => (contact.Parent as Outlook.MAPIFolder)?.StoreID),
                ItemType = "contact",
                MessageClass = SafeGet(() => contact.MessageClass),
                Subject = SafeGet(() => contact.CompanyName),
                Name = SafeGet(() => contact.FullName),
                Preview = includePreview
                    ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => contact.Body))
                    : null
            };
        }

        // Reason: rawItem is an Outlook item of a type this method could not identify - a PostItem,
        // JournalItem, DistListItem, or an item from a third-party add-in. The PIA models these as
        // unrelated COM classes with no common interface exposing MessageClass, Subject, FullName or
        // Name, so late binding is the only way to read them. SafeGet swallows the resulting
        // RuntimeBinderException when a given type does not have the member.
        dynamic untypedItem = rawItem;

        return new OutlookFolderItemInfo
        {
            ItemType = SafeGet(() => rawItem.GetType().Name),
            MessageClass = SafeGet(() => (string?)untypedItem.MessageClass),
            Subject = SafeGet(() => (string?)untypedItem.Subject),
            Name = SafeGet(() => (string?)untypedItem.FullName) ?? SafeGet(() => (string?)untypedItem.Name)
        };
    }

    private static string? TryGetDefaultRole(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            return "current";
        }

        return DefaultFolderAliases.ContainsKey(folder) ? folder.ToLowerInvariant() : null;
    }
}
