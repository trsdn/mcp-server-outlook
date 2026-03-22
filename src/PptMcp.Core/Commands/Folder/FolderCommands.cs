using System.Diagnostics.CodeAnalysis;
using PptMcp.Core.Commands.OutlookInterop;
using PptMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PptMcp.Core.Commands.Folder;

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
                    object? items = null;

                    try
                    {
                        folder = session.GetDefaultFolder(entry.Folder);
                        int? itemCount = null;
                        if (includeItemCounts)
                        {
                            items = folder.Items;
                            itemCount = ((dynamic)items).Count;
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
                        object? items = null;

                        try
                        {
                            childFolder = childFolders[index];
                            int? itemCount = null;
                            if (includeItemCounts)
                            {
                                items = childFolder.Items;
                                itemCount = ((dynamic)items).Count;
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
}
