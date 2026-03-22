using System.Diagnostics.CodeAnalysis;
using PptMcp.Core.Commands.OutlookInterop;
using PptMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PptMcp.Core.Commands.Application;

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
            ex => new OutlookApplicationStatusResult
            {
                Success = false,
                Connected = false,
                ErrorMessage = $"Failed to inspect Outlook application state: {ex.Message}"
            });
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
