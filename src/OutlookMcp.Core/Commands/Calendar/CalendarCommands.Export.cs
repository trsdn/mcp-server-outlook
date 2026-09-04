using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Calendar;

/// <summary>
/// Appointment export via <c>AppointmentItem.SaveAs</c> (#14). The destination rules and the format
/// table are shared with mail export in <see cref="ItemExportPlanner"/>; the only thing specific to
/// calendar is that <c>.ics</c> is the format this item type can actually produce.
/// </summary>
public partial class CalendarCommands
{
    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ItemExportResult Export(
        string filePath,
        string? format = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = true,
        bool overwrite = false)
    {
        if (!ItemExportPlanner.TryPlan(filePath, format, overwrite, out var plan, out var planError))
        {
            return planError!;
        }

        return OutlookInteropRunner.Execute(
            "OutlookCalendarExport",
            (application, session) =>
            {
                Outlook.AppointmentItem? appointment = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    appointment = ResolveAppointmentItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveAppointment,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (appointment == null)
                    {
                        return new ItemExportResult
                        {
                            Success = false,
                            ErrorMessage = entryId == null
                                ? "No appointment to export. Pass entryId, or open or select an appointment in Outlook."
                                : $"No appointment resolved for entry id '{entryId}'."
                        };
                    }

                    string? subject = SafeGet(() => appointment.Subject);
                    appointment.SaveAs(plan.FilePath, plan.Type);

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
                        EntryId = SafeGet(() => appointment.EntryID),
                        StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Message = $"Exported Outlook appointment to {plan.Format}."
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
                    OutlookInteropRunner.ReleaseComObject(ref appointment);
                }
            },
            ex => new ItemExportResult
            {
                Success = false,
                ErrorMessage = $"Failed to export the Outlook appointment: {ex.Message}"
            });
    }
}
