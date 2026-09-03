using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Calendar;

[ServiceCategory("calendar")]
[McpTool("calendar", Title = "Outlook Calendar Operations", Destructive = true, Category = "calendar",
    Description = "Inspect Outlook calendar items and manage appointments without opening a persistent session. "
    + "Use list to inspect the default Calendar folder or a specific Outlook folder path. "
    + "Use read to inspect an explicit appointment by entry id/store id or fall back to the active appointment inspector. "
    + "Use create-appointment, update-appointment, and delete-appointment to create, modify, and permanently delete Outlook calendar items — delete-appointment and update-appointment are destructive and cannot be undone. "
    + "Naming requiredAttendees or optionalAttendees (semicolon-separated) turns a new item into a meeting; it is saved to your own calendar and nobody is told until sendInvitation is true. "
    + "Attendees Outlook cannot resolve are reported and the meeting is not created, because an unresolved attendee never receives the invitation.")]
public interface ICalendarCommands
{
    [ServiceAction("list", Destructive = false)]
    CalendarListResult List(
        string? folder = null,
        string? start = null,
        string? endTime = null,
        int maxCount = 25,
        bool includeBodyPreview = false);

    [ServiceAction("read", Destructive = false)]
    CalendarItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = true);

    [ServiceAction("create-appointment", Destructive = true)]
    CalendarAppointmentResult CreateAppointment(
        string subject,
        string start,
        string endTime,
        string? location = null,
        string? body = null,
        bool allDay = false,
        bool display = false,
        string? requiredAttendees = null,
        string? optionalAttendees = null,
        bool sendInvitation = false);

    [ServiceAction("update-appointment", Destructive = true)]
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

    [ServiceAction("delete-appointment", Destructive = true)]
    CalendarMutationResult DeleteAppointment(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = false);
}
