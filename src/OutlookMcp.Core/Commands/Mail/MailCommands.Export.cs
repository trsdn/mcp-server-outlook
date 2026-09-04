using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// Mail item export via <c>MailItem.SaveAs</c> (#14). The destination rules, the format table and
/// the reason <c>msg</c> means <c>olMSGUnicode</c> all live in <see cref="ItemExportPlanner"/>,
/// which the calendar side shares.
/// </summary>
public partial class MailCommands
{
    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ItemExportResult Export(
        string filePath,
        string? format = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool overwrite = false)
    {
        if (!ItemExportPlanner.TryPlan(filePath, format, overwrite, out var plan, out var planError))
        {
            return planError!;
        }

        // A mail item cannot be an iCalendar entry. Outlook's own answer is "Value does not fall
        // within the expected range", which reads like a bug in the caller's arguments rather than
        // a statement about the item, so it is refused here with something actionable instead.
        if (plan.Format == "ics")
        {
            return new ItemExportResult
            {
                Success = false,
                ErrorMessage =
                    "A mail item cannot be exported as iCalendar (.ics). Use calendar export for "
                    + "appointments, or choose msg, txt, html, mht or rtf for mail."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailExport",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new ItemExportResult
                        {
                            Success = false,
                            ErrorMessage = entryId == null
                                ? "No mail item to export. Pass entryId, or open or select a message in Outlook."
                                : $"No mail item resolved for entry id '{entryId}'."
                        };
                    }

                    string? subject = SafeGet(() => mail.Subject);
                    mail.SaveAs(plan.FilePath, plan.Type);

                    var unwritten = ItemExportPlanner.VerifyWritten(plan);
                    if (unwritten != null)
                    {
                        return unwritten;
                    }

                    return new ItemExportResult
                    {
                        Success = true,
                        FilePath = plan.FilePath,
                        Format = plan.Format,
                        BytesWritten = new FileInfo(plan.FilePath).Length,
                        Overwritten = plan.Overwriting,
                        Subject = subject,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Message = $"Exported Outlook mail to {plan.Format}."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new ItemExportResult
            {
                Success = false,
                ErrorMessage = $"Failed to export the Outlook mail item: {ex.Message}"
            });
    }
}
