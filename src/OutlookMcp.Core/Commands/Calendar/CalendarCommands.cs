using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Calendar;

public class CalendarCommands : ICalendarCommands
{
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> FolderAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["calendar"] = Outlook.OlDefaultFolders.olFolderCalendar,
        ["current"] = Outlook.OlDefaultFolders.olFolderCalendar
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public CalendarListResult List(
        string? folder = null,
        string? start = null,
        string? endTime = null,
        int maxCount = 25,
        bool includeBodyPreview = false)
    {
        if (!TryParseRange(start, endTime, out DateTimeOffset? rangeStart, out DateTimeOffset? rangeEnd, out string? parseError))
        {
            return new CalendarListResult
            {
                Success = false,
                ErrorMessage = parseError
            };
        }

        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);

        return OutlookInteropRunner.Execute(
            "OutlookCalendarList",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? calendarFolder = null;
                Outlook.Items? items = null;

                try
                {
                    explorer = application.ActiveExplorer();
                    calendarFolder = ResolveCalendarFolder(application, session, folder, ref explorer);
                    if (calendarFolder == null)
                    {
                        return new CalendarListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = calendarFolder.Items;
                    TrySortItemsByStart(items);
                    int totalItemCount = SafeGetInt(() => items.Count);
                    int scanLimit = rangeStart.HasValue || rangeEnd.HasValue
                        ? totalItemCount
                        : Math.Clamp(boundedMaxCount * 10, 25, 500);

                    var result = new CalendarListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => calendarFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(calendarFolder),
                        Start = rangeStart,
                        End = rangeEnd,
                        TotalItemCount = totalItemCount
                    };

                    for (int index = 1, scanned = 0;
                         index <= totalItemCount && scanned < scanLimit && result.Appointments.Count < boundedMaxCount;
                         index++)
                    {
                        object? rawItem = null;
                        Outlook.AppointmentItem? appointment = null;

                        try
                        {
                            rawItem = items[index];
                            scanned++;
                            appointment = rawItem as Outlook.AppointmentItem;
                            if (appointment == null || !MatchesRange(appointment, rangeStart, rangeEnd))
                            {
                                continue;
                            }

                            result.Appointments.Add(CreateCalendarSummary(appointment, includeBodyPreview));
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref appointment);
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ReturnedCount = result.Appointments.Count;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref calendarFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new CalendarListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook calendar items: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public CalendarItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookCalendarRead",
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
                        return new CalendarItemResult
                        {
                            Success = true,
                            HasItem = false
                        };
                    }

                    return CreateCalendarItemResult(appointment);
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
            ex => new CalendarItemResult
            {
                Success = false,
                HasItem = false,
                ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                    ? $"Failed to inspect the active Outlook appointment: {ex.Message}"
                    : $"Failed to inspect the requested Outlook appointment: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public CalendarAppointmentResult CreateAppointment(
        string subject,
        string start,
        string endTime,
        string? location = null,
        string? body = null,
        bool allDay = false,
        bool display = false)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            return new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = "subject is required for calendar.create-appointment."
            };
        }

        if (!TryParseRange(start, endTime, out DateTimeOffset? parsedStart, out DateTimeOffset? parsedEnd, out string? parseError))
        {
            return new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = parseError
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookCalendarCreateAppointment",
            (application, session) =>
            {
                object? createdItem = null;
                Outlook.AppointmentItem? appointment = null;

                try
                {
                    createdItem = application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                    appointment = createdItem as Outlook.AppointmentItem;
                    if (appointment == null)
                    {
                        return new CalendarAppointmentResult
                        {
                            Success = false,
                            Saved = false,
                            Displayed = false,
                            ErrorMessage = "Outlook did not return an appointment item."
                        };
                    }

                    appointment.Subject = subject;
                    appointment.Start = parsedStart!.Value.LocalDateTime;
                    appointment.End = parsedEnd!.Value.LocalDateTime;
                    appointment.AllDayEvent = allDay;

                    if (!string.IsNullOrWhiteSpace(location))
                    {
                        appointment.Location = location;
                    }

                    if (body != null)
                    {
                        appointment.Body = body;
                    }

                    appointment.Save();
                    if (display)
                    {
                        appointment.Display(false);
                    }

                    return new CalendarAppointmentResult
                    {
                        Success = true,
                        Saved = true,
                        Displayed = display,
                        EntryId = SafeGet(() => appointment.EntryID),
                        StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => appointment.Subject),
                        Location = SafeGet(() => appointment.Location),
                        Start = SafeGetDateTimeOffset(() => appointment.Start),
                        End = SafeGetDateTimeOffset(() => appointment.End),
                        AllDay = SafeGetBool(() => appointment.AllDayEvent),
                        Message = "Created Outlook appointment."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref appointment);
                    OutlookInteropRunner.ReleaseComObject(ref createdItem);
                }
            },
            ex => new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = $"Failed to create the Outlook appointment: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public CalendarMutationResult UpdateAppointment(
        string? entryId = null,
        string? storeId = null,
        string? subject = null,
        string? start = null,
        string? endTime = null,
        string? location = null,
        string? body = null,
        bool? allDay = null,
        bool useActiveAppointment = false)
    {
        if (!TryParseRange(start, endTime, out DateTimeOffset? parsedStart, out DateTimeOffset? parsedEnd, out string? parseError))
        {
            return new CalendarMutationResult
            {
                Success = false,
                Updated = false,
                Deleted = false,
                ErrorMessage = parseError
            };
        }

        if (subject == null && start == null && endTime == null && location == null && body == null && allDay == null)
        {
            return new CalendarMutationResult
            {
                Success = false,
                Updated = false,
                Deleted = false,
                ErrorMessage = "At least one field must be supplied for calendar.update-appointment."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookCalendarUpdateAppointment",
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
                        return new CalendarMutationResult
                        {
                            Success = false,
                            Updated = false,
                            Deleted = false,
                            ErrorMessage = "Unable to resolve the Outlook appointment to update."
                        };
                    }

                    if (subject != null)
                    {
                        appointment.Subject = subject;
                    }

                    if (parsedStart.HasValue)
                    {
                        appointment.Start = parsedStart.Value.LocalDateTime;
                    }

                    if (parsedEnd.HasValue)
                    {
                        appointment.End = parsedEnd.Value.LocalDateTime;
                    }

                    if (location != null)
                    {
                        appointment.Location = location;
                    }

                    if (body != null)
                    {
                        appointment.Body = body;
                    }

                    if (allDay.HasValue)
                    {
                        appointment.AllDayEvent = allDay.Value;
                    }

                    appointment.Save();

                    return CreateCalendarMutationResult(
                        appointment,
                        updated: true,
                        deleted: false,
                        "Updated Outlook appointment.");
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
            ex => new CalendarMutationResult
            {
                Success = false,
                Updated = false,
                Deleted = false,
                ErrorMessage = $"Failed to update the Outlook appointment: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public CalendarMutationResult DeleteAppointment(
        string? entryId = null,
        string? storeId = null,
        bool useActiveAppointment = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookCalendarDeleteAppointment",
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
                        return new CalendarMutationResult
                        {
                            Success = false,
                            Updated = false,
                            Deleted = false,
                            ErrorMessage = "Unable to resolve the Outlook appointment to delete."
                        };
                    }

                    var result = CreateCalendarMutationResult(
                        appointment,
                        updated: false,
                        deleted: true,
                        "Deleted Outlook appointment.");

                    appointment.Delete();
                    return result;
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
            ex => new CalendarMutationResult
            {
                Success = false,
                Updated = false,
                Deleted = false,
                ErrorMessage = $"Failed to delete the Outlook appointment: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.AppointmentItem? ResolveAppointmentItem(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? entryId,
        string? storeId,
        bool useActiveAppointment,
        out Outlook.Inspector? inspector,
        out Outlook.Explorer? explorer,
        out Outlook.Selection? selection,
        out object? currentItem,
        out object? selectedItem,
        out object? resolvedItem)
    {
        inspector = null;
        explorer = null;
        selection = null;
        currentItem = null;
        selectedItem = null;
        resolvedItem = null;

        if (!string.IsNullOrWhiteSpace(entryId))
        {
            try
            {
                resolvedItem = session.GetItemFromID(entryId, string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
                return resolvedItem as Outlook.AppointmentItem;
            }
            catch
            {
                return null;
            }
        }

        if (!useActiveAppointment)
        {
            return null;
        }

        inspector = application.ActiveInspector();
        if (inspector != null)
        {
            currentItem = inspector.CurrentItem;
            if (currentItem is Outlook.AppointmentItem currentAppointment)
            {
                return currentAppointment;
            }
        }

        explorer = application.ActiveExplorer();
        if (explorer != null)
        {
            selection = explorer.Selection;
            if (selection != null && selection.Count > 0)
            {
                selectedItem = selection[1];
                if (selectedItem is Outlook.AppointmentItem selectedAppointment)
                {
                    return selectedAppointment;
                }
            }
        }

        return null;
    }

    private static bool TryParseRange(
        string? start,
        string? end,
        out DateTimeOffset? parsedStart,
        out DateTimeOffset? parsedEnd,
        out string? errorMessage)
    {
        parsedStart = null;
        parsedEnd = null;
        errorMessage = null;
        DateTimeOffset parsedStartValue = default;
        DateTimeOffset parsedEndValue = default;

        if (!string.IsNullOrWhiteSpace(start)
            && !DateTimeOffset.TryParse(start, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsedStartValue))
        {
            errorMessage = "start must be a valid ISO date/time value for Outlook calendar actions.";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(end)
            && !DateTimeOffset.TryParse(end, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsedEndValue))
        {
            errorMessage = "end must be a valid ISO date/time value for Outlook calendar actions.";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(start))
        {
            parsedStart = parsedStartValue;
        }

        if (!string.IsNullOrWhiteSpace(end))
        {
            parsedEnd = parsedEndValue;
        }

        if (parsedStart.HasValue && parsedEnd.HasValue && parsedEnd.Value < parsedStart.Value)
        {
            errorMessage = "end must be greater than or equal to start for Outlook calendar actions.";
            return false;
        }

        return true;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static CalendarItemResult CreateCalendarItemResult(Outlook.AppointmentItem appointment)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = appointment.Parent as Outlook.MAPIFolder;
            return new CalendarItemResult
            {
                Success = true,
                HasItem = true,
                EntryId = SafeGet(() => appointment.EntryID),
                StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                Subject = SafeGet(() => appointment.Subject),
                Location = SafeGet(() => appointment.Location),
                Organizer = SafeGet(() => appointment.Organizer),
                BodyPreview = OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => appointment.Body)),
                FolderPath = OutlookInteropRunner.GetFolderPath(parentFolder),
                Start = SafeGetDateTimeOffset(() => appointment.Start),
                End = SafeGetDateTimeOffset(() => appointment.End),
                AllDay = SafeGetBool(() => appointment.AllDayEvent),
                ReminderSet = SafeGetBool(() => appointment.ReminderSet),
                BusyStatus = SafeGetInt(() => (int)appointment.BusyStatus)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static CalendarSummaryInfo CreateCalendarSummary(Outlook.AppointmentItem appointment, bool includeBodyPreview)
    {
        return new CalendarSummaryInfo
        {
            EntryId = SafeGet(() => appointment.EntryID),
            StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
            Subject = SafeGet(() => appointment.Subject),
            Location = SafeGet(() => appointment.Location),
            Organizer = SafeGet(() => appointment.Organizer),
            BodyPreview = includeBodyPreview
                ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => appointment.Body))
                : null,
            Start = SafeGetDateTimeOffset(() => appointment.Start),
            End = SafeGetDateTimeOffset(() => appointment.End),
            AllDay = SafeGetBool(() => appointment.AllDayEvent),
            ReminderSet = SafeGetBool(() => appointment.ReminderSet),
            BusyStatus = SafeGetInt(() => (int)appointment.BusyStatus)
        };
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static CalendarMutationResult CreateCalendarMutationResult(
        Outlook.AppointmentItem appointment,
        bool updated,
        bool deleted,
        string message)
    {
        return new CalendarMutationResult
        {
            Success = true,
            EntryId = SafeGet(() => appointment.EntryID),
            StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
            Subject = SafeGet(() => appointment.Subject),
            Location = SafeGet(() => appointment.Location),
            Start = SafeGetDateTimeOffset(() => appointment.Start),
            End = SafeGetDateTimeOffset(() => appointment.End),
            AllDay = SafeGetBool(() => appointment.AllDayEvent),
            Updated = updated,
            Deleted = deleted,
            Message = message
        };
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool MatchesRange(Outlook.AppointmentItem appointment, DateTimeOffset? rangeStart, DateTimeOffset? rangeEnd)
    {
        DateTimeOffset? start = SafeGetDateTimeOffset(() => appointment.Start);
        DateTimeOffset? end = SafeGetDateTimeOffset(() => appointment.End);
        if (!start.HasValue || !end.HasValue)
        {
            return false;
        }

        if (rangeStart.HasValue && end.Value < rangeStart.Value)
        {
            return false;
        }

        if (rangeEnd.HasValue && start.Value > rangeEnd.Value)
        {
            return false;
        }

        return true;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void TrySortItemsByStart(Outlook.Items items)
    {
        try
        {
            items.Sort("[Start]", false);
        }
        catch
        {
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? ResolveCalendarFolder(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? folder,
        ref Outlook.Explorer? explorer)
        => OutlookInteropRunner.ResolveFolder(
            application,
            session,
            string.IsNullOrWhiteSpace(folder) ? "calendar" : folder,
            FolderAliases,
            ref explorer);

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

    private static DateTimeOffset? SafeGetDateTimeOffset(Func<DateTime> getter)
    {
        try
        {
            DateTime value = getter();
            return value == default ? null : new DateTimeOffset(value);
        }
        catch
        {
            return null;
        }
    }

    private static string BuildUnknownFolderMessage(string? folder)
    {
        const string supportedFolders = "calendar, current, or an Outlook calendar folder path";
        return string.IsNullOrWhiteSpace(folder)
            ? $"Could not resolve the Outlook calendar folder. Supported folder values: {supportedFolders}."
            : $"Unsupported Outlook calendar folder '{folder}'. Supported folder values: {supportedFolders}.";
    }
}
