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
    + "Attendees Outlook cannot resolve are reported and the meeting is not created, because an unresolved attendee never receives the invitation. "
    + "Use get-free-busy to ask when one or more people are available before proposing a time; it returns both Outlook's raw slot string and the busy periods decoded from it. "
    + "Recurring series: pass recurrenceType (daily, weekly, monthly or yearly) to create-appointment to make a series, "
    + "with recurrenceInterval, recurrenceDaysOfWeek (semicolon-separated day names, weekly only) and either "
    + "recurrenceCount or recurrenceEndDate to bound it. list only expands a series into its individual occurrences "
    + "when both start and endTime are given, because a series with no end date has infinitely many; it reports "
    + "recurringExpanded so a caller can tell. Never conclude somebody is free from a listing whose recurringExpanded "
    + "is false - it contains series masters only, so a weekly meeting is missing from every date but its first. "
    + "Every occurrence of a series carries the series' own entry id, so update-appointment and delete-appointment "
    + "against that entry id change or cancel the WHOLE series. To change or cancel one instance, pass occurrenceDate "
    + "with the date of that instance; a bare date takes its time of day from the series. Naming occurrenceDate on an "
    + "item that is not recurring is refused rather than silently ignored, and the response reports scope as series or "
    + "occurrence so a caller can confirm what was actually touched.")]
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
        string? resourceAttendees = null,
        bool sendInvitation = false,
        string? recurrenceType = null,
        int recurrenceInterval = 1,
        string? recurrenceDaysOfWeek = null,
        int? recurrenceCount = null,
        string? recurrenceEndDate = null);

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
        bool useActiveAppointment = false,
        string? occurrenceDate = null);

    [ServiceAction("delete-appointment", Destructive = true)]
    CalendarMutationResult DeleteAppointment(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = false,
        string? occurrenceDate = null);

    [ServiceAction("get-free-busy", Destructive = false)]
    CalendarFreeBusyResult GetFreeBusy(
        string attendees,
        string? start = null,
        int days = 7,
        int intervalMinutes = 30);

    /// <summary>
    /// Saves an appointment to disk with <c>AppointmentItem.SaveAs</c>.
    ///
    /// <para>
    /// This is the half of item export that can produce iCalendar: a mail item asked for
    /// <c>.ics</c> is refused, because Outlook answers it with "Value does not fall within the
    /// expected range". <paramref name="filePath"/> must be absolute - Outlook resolves a relative
    /// path against its own working directory - and an existing file is never replaced unless
    /// <paramref name="overwrite"/> is set.
    /// </para>
    /// </summary>
    [ServiceAction("export", Destructive = false)]
    ItemExportResult Export(
        string filePath,
        string? format = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = true,
        bool overwrite = false);
}
