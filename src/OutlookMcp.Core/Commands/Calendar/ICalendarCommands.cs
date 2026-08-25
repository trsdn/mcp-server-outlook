using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Calendar;

[ServiceCategory("calendar")]
[NoSession]
[McpTool("calendar", Title = "Outlook Calendar Operations", Destructive = false, Category = "calendar",
    Description = "Inspect Outlook calendar items and create safe appointments without opening a persistent session. "
    + "Use list to inspect the default Calendar folder or a specific Outlook folder path. "
    + "Use read to inspect an explicit appointment by entry id/store id or fall back to the active appointment inspector. "
    + "Use create-appointment, update-appointment, and delete-appointment to manage Outlook calendar items safely.")]
public interface ICalendarCommands
{
    [ServiceAction("list")]
    CalendarListResult List(
        string? folder = null,
        string? start = null,
        string? endTime = null,
        int maxCount = 25,
        bool includeBodyPreview = false);

    [ServiceAction("read")]
    CalendarItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = true);

    [ServiceAction("create-appointment")]
    CalendarAppointmentResult CreateAppointment(
        string subject,
        string start,
        string endTime,
        string? location = null,
        string? body = null,
        bool allDay = false,
        bool display = false);

    [ServiceAction("update-appointment")]
    CalendarMutationResult UpdateAppointment(
        string? entryId = null,
        string? storeId = null,
        string? subject = null,
        string? start = null,
        string? endTime = null,
        string? location = null,
        string? body = null,
        bool? allDay = null,
        bool useActiveAppointment = false);

    [ServiceAction("delete-appointment")]
    CalendarMutationResult DeleteAppointment(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = false);
}
