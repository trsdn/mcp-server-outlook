using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Application;

public class ApplicationCommands : IApplicationCommands
{
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookApplicationStatusResult GetStatus(bool includeActiveContext = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookApplicationGetStatus",
            (application, session) =>
            {
                Outlook.Explorer? activeExplorer = null;
                object? currentFolderObject = null;
                Outlook.MAPIFolder? currentFolder = null;

                try
                {
                    activeExplorer = application.ActiveExplorer();
                    if (activeExplorer != null)
                    {
                        currentFolderObject = activeExplorer.CurrentFolder;
                        currentFolder = currentFolderObject as Outlook.MAPIFolder;
                    }

                    return new OutlookApplicationStatusResult
                    {
                        Success = true,
                        Connected = true,
                        Version = application.Version ?? string.Empty,
                        ExplorerCount = application.Explorers.Count,
                        InspectorCount = application.Inspectors.Count,
                        StoreCount = session.Folders.Count,
                        OutlookFlavor = "classic-desktop",
                        ProcessElevated = OutlookInstallationDetector.IsCurrentProcessElevated(),
                        CurrentFolderName = includeActiveContext ? currentFolder?.Name : null,
                        CurrentFolderPath = includeActiveContext ? OutlookInteropRunner.GetFolderPath(currentFolder) : null,
                        HasActiveMailSelection = includeActiveContext && HasActiveMailSelection(activeExplorer)
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref currentFolder);
                    OutlookInteropRunner.ReleaseComObject(ref currentFolderObject);
                    OutlookInteropRunner.ReleaseComObject(ref activeExplorer);
                }
            },
            ex =>
            {
                OutlookFlavor flavor = OutlookInstallationDetector.DetectFlavor();
                return new OutlookApplicationStatusResult
                {
                    Success = false,
                    Connected = false,
                    OutlookFlavor = flavor.ToString(),
                    ProcessElevated = OutlookInstallationDetector.IsCurrentProcessElevated(),
                    ErrorMessage = flavor == OutlookFlavor.ClassicDesktop
                        ? $"Classic Outlook is installed but could not be reached (installed but not running, or running at a different integrity level): {ex.Message}"
                        : $"{OutlookInstallationDetector.BuildUnavailableMessage(flavor)} ({ex.Message})"
                };
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookExplorerContextResult GetActiveExplorer()
    {
        return OutlookInteropRunner.Execute(
            "OutlookApplicationGetActiveExplorer",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                object? currentFolderObject = null;
                Outlook.MAPIFolder? currentFolder = null;
                Outlook.Selection? selection = null;
                object? selectedItem = null;

                try
                {
                    explorer = application.ActiveExplorer();
                    if (explorer == null)
                    {
                        return new OutlookExplorerContextResult { Success = true, HasExplorer = false };
                    }

                    currentFolderObject = explorer.CurrentFolder;
                    currentFolder = currentFolderObject as Outlook.MAPIFolder;

                    selection = explorer.Selection;
                    int selectionCount = selection?.Count ?? 0;
                    if (selectionCount > 0)
                    {
                        selectedItem = selection![1];
                    }

                    var result = new OutlookExplorerContextResult
                    {
                        Success = true,
                        HasExplorer = true,
                        CurrentFolderName = currentFolder?.Name,
                        CurrentFolderPath = OutlookInteropRunner.GetFolderPath(currentFolder),
                        SelectionCount = selectionCount,
                        HasMailSelection = selectedItem is Outlook.MailItem
                    };

                    DescribeItem(
                        selectedItem,
                        out string? itemType,
                        out string? messageClass,
                        out string? subject,
                        out string? entryId,
                        out _);

                    result.SelectedItemType = itemType;
                    result.SelectedItemMessageClass = messageClass;
                    result.SelectedItemSubject = subject;
                    result.SelectedItemEntryId = entryId;

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref currentFolder);
                    OutlookInteropRunner.ReleaseComObject(ref currentFolderObject);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookExplorerContextResult
            {
                Success = false,
                HasExplorer = false,
                ErrorMessage = $"Failed to inspect the active Outlook explorer: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookInspectorContextResult GetActiveInspector()
    {
        return OutlookInteropRunner.Execute(
            "OutlookApplicationGetActiveInspector",
            (application, session) =>
            {
                Outlook.Inspector? inspector = null;
                object? currentItem = null;

                try
                {
                    inspector = application.ActiveInspector();
                    if (inspector == null)
                    {
                        return new OutlookInspectorContextResult { Success = true, HasInspector = false };
                    }

                    currentItem = inspector.CurrentItem;

                    DescribeItem(
                        currentItem,
                        out string? itemType,
                        out string? messageClass,
                        out string? subject,
                        out string? entryId,
                        out bool isSaved);

                    return new OutlookInspectorContextResult
                    {
                        Success = true,
                        HasInspector = true,
                        ItemType = itemType,
                        MessageClass = messageClass,
                        Subject = subject,
                        EntryId = entryId,
                        StoreId = GetParentStoreId(currentItem),
                        IsSaved = isSaved,
                        CurrentFolderPath = GetParentFolderPath(currentItem),
                        Caption = inspector.Caption
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                }
            },
            ex => new OutlookInspectorContextResult
            {
                Success = false,
                HasInspector = false,
                ErrorMessage = $"Failed to inspect the active Outlook inspector: {ex.Message}"
            });
    }

    /// <summary>
    /// Describes an Outlook item without asking it for its runtime type. The wrapper around a COM
    /// item reports <c>__ComObject</c> from <c>GetType().Name</c>, which is what the earlier draft
    /// of this code surfaced to callers, so the kind is taken from the typed item instead.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void DescribeItem(
        object? item,
        out string? itemType,
        out string? messageClass,
        out string? subject,
        out string? entryId,
        out bool isSaved)
    {
        itemType = null;
        messageClass = null;
        subject = null;
        entryId = null;
        isSaved = false;

        switch (item)
        {
            case null:
                return;

            case Outlook.MailItem mail:
                itemType = "mail";
                messageClass = mail.MessageClass;
                subject = mail.Subject;
                entryId = mail.EntryID;
                isSaved = mail.Saved;
                break;

            case Outlook.AppointmentItem appointment:
                itemType = "appointment";
                messageClass = appointment.MessageClass;
                subject = appointment.Subject;
                entryId = appointment.EntryID;
                isSaved = appointment.Saved;
                break;

            case Outlook.MeetingItem meeting:
                itemType = "meetingResponse";
                messageClass = meeting.MessageClass;
                subject = meeting.Subject;
                entryId = meeting.EntryID;
                isSaved = meeting.Saved;
                break;

            case Outlook.ContactItem contact:
                itemType = "contact";
                messageClass = contact.MessageClass;
                subject = contact.FullName;
                entryId = contact.EntryID;
                isSaved = contact.Saved;
                break;

            case Outlook.DistListItem distributionList:
                itemType = "distributionList";
                messageClass = distributionList.MessageClass;
                subject = distributionList.DLName;
                entryId = distributionList.EntryID;
                isSaved = distributionList.Saved;
                break;

            case Outlook.TaskItem task:
                itemType = "task";
                messageClass = task.MessageClass;
                subject = task.Subject;
                entryId = task.EntryID;
                isSaved = task.Saved;
                break;

            case Outlook.NoteItem note:
                itemType = "note";
                messageClass = note.MessageClass;
                subject = note.Subject;
                entryId = note.EntryID;
                isSaved = note.Saved;
                break;

            default:
                // A kind this build does not model. Say so plainly rather than guessing: an agent
                // can still act on the message class, and "unknown" is honest.
                itemType = "unknown";
                break;
        }

        if (string.IsNullOrEmpty(subject))
        {
            subject = null;
        }

        // An item that has never been saved has no entry id. The PIA returns null for it; a raw
        // PowerShell probe renders that as an empty string, which is the same thing seen through a
        // lossy lens. Normalise both to null so the field is simply absent, rather than handing the
        // caller a handle that resolves to nothing.
        if (string.IsNullOrEmpty(entryId))
        {
            entryId = null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? GetParentFolderPath(object? item)
    {
        Outlook.MAPIFolder? parent = null;

        try
        {
            parent = GetParentFolder(item);
            return OutlookInteropRunner.GetFolderPath(parent);
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parent);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? GetParentStoreId(object? item)
    {
        Outlook.MAPIFolder? parent = null;

        try
        {
            parent = GetParentFolder(item);
            return parent?.StoreID;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parent);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? GetParentFolder(object? item)
    {
        return item switch
        {
            Outlook.MailItem mail => mail.Parent as Outlook.MAPIFolder,
            Outlook.AppointmentItem appointment => appointment.Parent as Outlook.MAPIFolder,
            Outlook.MeetingItem meeting => meeting.Parent as Outlook.MAPIFolder,
            Outlook.ContactItem contact => contact.Parent as Outlook.MAPIFolder,
            Outlook.DistListItem distributionList => distributionList.Parent as Outlook.MAPIFolder,
            Outlook.TaskItem task => task.Parent as Outlook.MAPIFolder,
            Outlook.NoteItem note => note.Parent as Outlook.MAPIFolder,
            _ => null
        };
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool HasActiveMailSelection(Outlook.Explorer? explorer)
    {
        Outlook.Selection? selection = null;
        object? selectedItem = null;

        try
        {
            if (explorer == null)
            {
                return false;
            }

            selection = explorer.Selection;
            if (selection == null || selection.Count < 1)
            {
                return false;
            }

            selectedItem = selection[1];
            return selectedItem is Outlook.MailItem;
        }
        catch
        {
            return false;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref selectedItem);
            OutlookInteropRunner.ReleaseComObject(ref selection);
        }
    }
}
