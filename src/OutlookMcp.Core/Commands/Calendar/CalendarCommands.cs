using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Calendar;

public partial class CalendarCommands : ICalendarCommands
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

                    // Expansion needs both bounds. Outlook enumerates a series occurrence by
                    // occurrence, so an endless series over an open-ended range never terminates.
                    bool canExpand = rangeStart.HasValue && rangeEnd.HasValue;

                    var result = new CalendarListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => calendarFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(calendarFolder),
                        Start = rangeStart,
                        End = rangeEnd,
                        TotalItemCount = totalItemCount,
                        RecurringExpanded = false
                    };

                    if (canExpand)
                    {
                        Outlook.Items? expanded = null;

                        try
                        {
                            expanded = TryExpandRecurrences(items, rangeStart!.Value, rangeEnd!.Value);

                            if (expanded != null)
                            {
                                result.RecurringExpanded = true;
                                result.Message = "Recurring series were expanded into their individual occurrences.";
                                CollectExpandedAppointments(
                                    expanded,
                                    rangeStart,
                                    rangeEnd,
                                    boundedMaxCount,
                                    includeBodyPreview,
                                    result);
                                result.ReturnedCount = result.Appointments.Count;
                                return result;
                            }

                            // Outlook refused the expansion. Saying so is the point: a caller that
                            // believes a listing is complete will report somebody as free when they
                            // are not.
                            result.Message = "Outlook would not expand recurring series for this folder, so occurrences "
                                + "of a recurring meeting are missing from every date but its first.";
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref expanded);
                        }
                    }
                    else
                    {
                        result.Message = "Recurring series were not expanded - that needs both start and endTime, "
                            + "because a series with no end date has infinitely many occurrences. This list contains "
                            + "series masters only, so a recurring meeting is missing from every date but its first.";
                    }

                    int scanLimit = rangeStart.HasValue || rangeEnd.HasValue
                        ? totalItemCount
                        : Math.Clamp(boundedMaxCount * 10, 25, 500);

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
        bool display = false,
        string? requiredAttendees = null,
        string? optionalAttendees = null,
        bool sendInvitation = false,
        string? recurrenceType = null,
        int recurrenceInterval = 1,
        string? recurrenceDaysOfWeek = null,
        int? recurrenceCount = null,
        string? recurrenceEndDate = null)
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

        bool wantsMeeting = !string.IsNullOrWhiteSpace(requiredAttendees) || !string.IsNullOrWhiteSpace(optionalAttendees);

        if (sendInvitation && !wantsMeeting)
        {
            return new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = "sendInvitation was requested but no attendees were named, so there is nobody to invite. "
                    + "Pass requiredAttendees or optionalAttendees."
            };
        }

        bool wantsRecurrence = !string.IsNullOrWhiteSpace(recurrenceType);
        Outlook.OlRecurrenceType parsedRecurrence = Outlook.OlRecurrenceType.olRecursDaily;

        // Validated before anything is written. Creating a plain appointment because the pattern was
        // unusable, and reporting success, would leave the caller believing they had made a series.
        if (wantsRecurrence && !TryParseRecurrenceType(recurrenceType, out parsedRecurrence))

        {
            return new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = $"'{recurrenceType}' is not a recurrence type calendar.create-appointment understands. "
                    + "Use daily, weekly, monthly or yearly."
            };
        }

        if (wantsRecurrence && recurrenceInterval < 1)
        {
            return new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = $"recurrenceInterval must be at least 1, got {recurrenceInterval}."
            };
        }

        if (wantsRecurrence && recurrenceCount.HasValue && !string.IsNullOrWhiteSpace(recurrenceEndDate))
        {
            return new CalendarAppointmentResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = "recurrenceCount and recurrenceEndDate both bound the series, and Outlook keeps only "
                    + "one of them. Pass whichever you actually mean, not both."
            };
        }

        DateTimeOffset? parsedRecurrenceEnd = null;

        if (wantsRecurrence && !string.IsNullOrWhiteSpace(recurrenceEndDate))
        {
            if (!TryParseRange(recurrenceEndDate, null, out parsedRecurrenceEnd, out _, out string? endParseError))
            {
                return new CalendarAppointmentResult
                {
                    Success = false,
                    Saved = false,
                    Displayed = false,
                    ErrorMessage = $"recurrenceEndDate could not be read: {endParseError}"
                };
            }
        }

        var parsedDays = new List<DayOfWeek>();

        if (wantsRecurrence && !string.IsNullOrWhiteSpace(recurrenceDaysOfWeek))
        {
            foreach (string token in recurrenceDaysOfWeek.Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (!TryParseDayOfWeek(token, out DayOfWeek day))
                {
                    return new CalendarAppointmentResult
                    {
                        Success = false,
                        Saved = false,
                        Displayed = false,
                        ErrorMessage = $"'{token}' is not a day name recurrenceDaysOfWeek understands. "
                            + "Use full English day names such as monday;thursday."
                    };
                }

                parsedDays.Add(day);
            }
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

                    var attendees = new List<MeetingAttendeeInfo>();

                    if (wantsMeeting)
                    {
                        appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;

                        if (!string.IsNullOrWhiteSpace(requiredAttendees))
                        {
                            appointment.RequiredAttendees = requiredAttendees;
                        }

                        if (!string.IsNullOrWhiteSpace(optionalAttendees))
                        {
                            appointment.OptionalAttendees = optionalAttendees;
                        }

                        attendees = ReadAttendees(appointment, resolveFirst: true);
                        List<string> unresolved = attendees
                            .Where(a => !a.Resolved)
                            .Select(a => a.Name ?? a.Address ?? "(unnamed)")
                            .ToList();

                        if (unresolved.Count > 0)
                        {
                            // Nothing has been saved yet, so there is no stray item to clean up.
                            // Saving anyway would report success for a meeting that can never reach
                            // the people the caller named.
                            return new CalendarAppointmentResult
                            {
                                Success = false,
                                Saved = false,
                                Displayed = false,
                                IsMeeting = true,
                                InvitationSent = false,
                                Attendees = attendees,
                                UnresolvedAttendees = unresolved,
                                ErrorMessage = "Outlook could not resolve these attendees, so no meeting was created: "
                                    + string.Join(", ", unresolved)
                                    + ". Use a full SMTP address or a name that exists in the address book."
                            };
                        }
                    }

                    appointment.Save();

                    RecurrencePatternInfo? recurrenceInfo = null;

                    if (wantsRecurrence)
                    {
                        Outlook.RecurrencePattern? pattern = null;

                        try
                        {
                            pattern = appointment.GetRecurrencePattern();
                            pattern.RecurrenceType = parsedRecurrence;
                            pattern.Interval = recurrenceInterval;

                            if (parsedRecurrence == Outlook.OlRecurrenceType.olRecursWeekly)
                            {
                                // A weekly pattern with no day named repeats on the start day. Guessing
                                // any other day would put the series somewhere the caller never asked for.
                                List<DayOfWeek> days = parsedDays.Count > 0
                                    ? parsedDays
                                    : [parsedStart!.Value.LocalDateTime.DayOfWeek];

                                pattern.DayOfWeekMask = ToDayOfWeekMask(days);
                            }

                            pattern.PatternStartDate = parsedStart!.Value.LocalDateTime.Date;

                            if (recurrenceCount.HasValue)
                            {
                                pattern.Occurrences = recurrenceCount.Value;
                            }
                            else if (parsedRecurrenceEnd.HasValue)
                            {
                                pattern.PatternEndDate = parsedRecurrenceEnd.Value.LocalDateTime.Date;
                            }
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref pattern);
                        }

                        appointment.Save();

                        // Microsoft's documented guidance: release the pattern and re-fetch before
                        // reading it back. A pattern held across a Save is a known source of
                        // corruption, and reading through a stale one is how a series quietly loses
                        // its exceptions.
                        recurrenceInfo = ReadRecurrence(appointment);
                    }

                    bool invitationSent = false;

                    if (sendInvitation)
                    {
                        appointment.Send();
                        invitationSent = true;
                    }

                    if (display)
                    {
                        appointment.Display(false);
                    }

                    return new CalendarAppointmentResult
                    {
                        Success = true,
                        Saved = true,
                        Displayed = display,
                        IsMeeting = wantsMeeting,
                        InvitationSent = invitationSent,
                        Attendees = attendees,
                        EntryId = SafeGet(() => appointment.EntryID),
                        StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => appointment.Subject),
                        Location = SafeGet(() => appointment.Location),
                        Start = SafeGetDateTimeOffset(() => appointment.Start),
                        End = SafeGetDateTimeOffset(() => appointment.End),
                        AllDay = SafeGetBool(() => appointment.AllDayEvent),
                        IsRecurring = recurrenceInfo != null,
                        Recurrence = recurrenceInfo,
                        Message = wantsMeeting
                            ? invitationSent
                                ? "Created Outlook meeting and sent the invitation."
                                : "Created Outlook meeting. No invitation was sent - the attendees have not been told. Pass sendInvitation to invite them."
                            : "Created Outlook appointment."
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
        bool useActiveAppointment = false,
        string? occurrenceDate = null)
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
                Outlook.AppointmentItem? occurrence = null;

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

                    // The target is the series unless the caller named an instance. Resolving the
                    // occurrence first means a bad occurrenceDate cannot leave a half-applied edit
                    // on the series behind it.
                    Outlook.AppointmentItem target = appointment;
                    string scope = "series";

                    if (!string.IsNullOrWhiteSpace(occurrenceDate))
                    {
                        if (!TryResolveOccurrence(appointment, occurrenceDate!, out occurrence, out string? occurrenceError)
                            || occurrence == null)
                        {
                            return new CalendarMutationResult
                            {
                                Success = false,
                                Updated = false,
                                Deleted = false,
                                ErrorMessage = occurrenceError
                            };
                        }

                        target = occurrence;
                        scope = "occurrence";
                    }

                    if (subject != null)
                    {
                        target.Subject = subject;
                    }

                    if (parsedStart.HasValue)
                    {
                        target.Start = parsedStart.Value.LocalDateTime;
                    }

                    if (parsedEnd.HasValue)
                    {
                        target.End = parsedEnd.Value.LocalDateTime;
                    }

                    if (location != null)
                    {
                        target.Location = location;
                    }

                    if (body != null)
                    {
                        target.Body = body;
                    }

                    if (allDay.HasValue)
                    {
                        target.AllDayEvent = allDay.Value;
                    }

                    target.Save();

                    return CreateCalendarMutationResult(
                        target,
                        updated: true,
                        deleted: false,
                        scope == "occurrence"
                            ? "Updated one occurrence of the Outlook series; the rest of the series is unchanged."
                            : "Updated Outlook appointment.",
                        scope);
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref occurrence);
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
        bool useActiveAppointment = false,
        string? occurrenceDate = null)
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
                Outlook.AppointmentItem? occurrence = null;

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

                    Outlook.AppointmentItem target = appointment;
                    string scope = "series";

                    if (!string.IsNullOrWhiteSpace(occurrenceDate))
                    {
                        if (!TryResolveOccurrence(appointment, occurrenceDate!, out occurrence, out string? occurrenceError)
                            || occurrence == null)
                        {
                            return new CalendarMutationResult
                            {
                                Success = false,
                                Updated = false,
                                Deleted = false,
                                ErrorMessage = occurrenceError
                            };
                        }

                        target = occurrence;
                        scope = "occurrence";
                    }

                    var result = CreateCalendarMutationResult(
                        target,
                        updated: false,
                        deleted: true,
                        scope == "occurrence"
                            ? "Cancelled one occurrence of the Outlook series; the rest of the series is intact."
                            : "Deleted Outlook appointment.",
                        scope);

                    target.Delete();
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref occurrence);
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

    /// <summary>
    /// Asks Outlook when the named people are available.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <c>Recipient.FreeBusy</c> returns one character per interval - <c>0</c> free, <c>1</c>
    /// tentative, <c>2</c> busy, <c>3</c> out of office, <c>4</c> working elsewhere - which is
    /// compact but useless to read directly. Both forms are returned: the raw string, and the
    /// non-free stretches decoded into timestamps.
    /// </para>
    /// <para>
    /// Outlook decides for itself how far ahead it publishes and ignores the requested length, so the
    /// string is trimmed to the window asked for, and <c>end</c> is pulled in when Outlook returned
    /// less than that. Padding it out would invent free time nobody looked up.
    /// </para>
    /// <para>
    /// An unresolvable attendee fails the whole call. Outlook returns an all-free string for a
    /// recipient it never looked up, so treating that as an answer would propose meetings on top of a
    /// calendar nobody ever read.
    /// </para>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public CalendarFreeBusyResult GetFreeBusy(
        string attendees,
        string? start = null,
        int days = 7,
        int intervalMinutes = 30)
    {
        string[] names = SplitAttendees(attendees);

        if (names.Length == 0)
        {
            return new CalendarFreeBusyResult
            {
                Success = false,
                ErrorMessage = "attendees is required for calendar.get-free-busy. "
                    + "Pass one or more names or SMTP addresses separated by semicolons."
            };
        }

        if (days < 1)
        {
            return new CalendarFreeBusyResult
            {
                Success = false,
                ErrorMessage = "days must be at least 1 for calendar.get-free-busy."
            };
        }

        // Outlook rejects an interval that does not divide the day evenly.
        if (intervalMinutes < 1 || 1440 % intervalMinutes != 0)
        {
            return new CalendarFreeBusyResult
            {
                Success = false,
                ErrorMessage = $"intervalMinutes must divide 1440 evenly (Outlook works in whole days); {intervalMinutes} does not. "
                    + "Try 5, 10, 15, 30, 60 or 120."
            };
        }

        if (!TryParseRange(start, null, out DateTimeOffset? parsedStart, out _, out string? parseError))
        {
            return new CalendarFreeBusyResult
            {
                Success = false,
                ErrorMessage = parseError
            };
        }

        // Outlook always starts the slot string at midnight of the requested day, so anchoring the
        // window there is the only reading of the result that lines up with the data.
        DateTimeOffset requested = parsedStart ?? DateTimeOffset.Now;
        DateTimeOffset windowStart = new(requested.Date, requested.Offset);
        DateTimeOffset windowEnd = windowStart.AddDays(days);

        return OutlookInteropRunner.Execute(
            "OutlookCalendarGetFreeBusy",
            (application, session) =>
            {
                object? probeItem = null;
                Outlook.AppointmentItem? probe = null;
                Outlook.Recipients? recipients = null;

                try
                {
                    // A Recipients collection has to belong to an item. This one exists only to own
                    // it and is never saved, so nothing reaches the calendar.
                    probeItem = application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                    probe = probeItem as Outlook.AppointmentItem;
                    recipients = probe?.Recipients;

                    if (recipients == null)
                    {
                        return new CalendarFreeBusyResult
                        {
                            Success = false,
                            ErrorMessage = "Outlook did not provide a recipients collection for the free/busy lookup."
                        };
                    }

                    foreach (string name in names)
                    {
                        Outlook.Recipient? added = null;

                        try
                        {
                            added = recipients.Add(name);
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref added);
                        }
                    }

                    _ = recipients.ResolveAll();

                    var people = new List<FreeBusyPersonInfo>();
                    var unresolved = new List<string>();
                    int count = recipients.Count;
                    int requestedSlots = days * (1440 / intervalMinutes);
                    int coveredSlots = requestedSlots;

                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Recipient? recipient = null;

                        try
                        {
                            recipient = recipients[index];
                            bool resolved = SafeGetBool(() => recipient.Resolved);
                            string? name = SafeGet(() => recipient.Name);

                            if (!resolved)
                            {
                                unresolved.Add(name ?? "(unnamed)");
                                people.Add(new FreeBusyPersonInfo
                                {
                                    Name = name,
                                    Address = SafeGet(() => recipient.Address),
                                    Resolved = false
                                });
                                continue;
                            }

                            string? raw = SafeGet(() =>
                                recipient.FreeBusy(windowStart.LocalDateTime, intervalMinutes, true));

                            // Outlook decides for itself how far ahead it publishes and ignores the
                            // requested length, so the string is trimmed to the window that was
                            // asked for - and the window is shortened when Outlook returned less.
                            // Padding it out would invent free time nobody looked up.
                            string? availability = raw is null
                                ? null
                                : raw.Length > requestedSlots ? raw[..requestedSlots] : raw;

                            if (availability != null && availability.Length < coveredSlots)
                            {
                                coveredSlots = availability.Length;
                            }

                            people.Add(new FreeBusyPersonInfo
                            {
                                Name = name,
                                Address = SafeGet(() => recipient.Address),
                                Resolved = true,
                                Availability = availability,
                                BusyPeriods = DecodeBusyPeriods(availability, windowStart, intervalMinutes)
                            });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref recipient);
                        }
                    }

                    if (unresolved.Count > 0)
                    {
                        return new CalendarFreeBusyResult
                        {
                            Success = false,
                            Start = windowStart,
                            End = windowEnd,
                            IntervalMinutes = intervalMinutes,
                            People = people,
                            UnresolvedAttendees = unresolved,
                            ErrorMessage = "Outlook could not resolve these attendees, so their availability is unknown: "
                                + string.Join(", ", unresolved)
                                + ". Use a full SMTP address or a name that exists in the address book."
                        };
                    }

                    DateTimeOffset coveredEnd = windowStart.AddMinutes((double)coveredSlots * intervalMinutes);

                    return new CalendarFreeBusyResult
                    {
                        Success = true,
                        Start = windowStart,
                        End = coveredEnd,
                        IntervalMinutes = intervalMinutes,
                        People = people,
                        Message = coveredSlots < requestedSlots
                            ? $"Outlook published availability only as far as {coveredEnd:yyyy-MM-dd HH:mm}, short of the {days} day(s) requested. "
                                + "Nothing is known about the remainder - do not treat it as free."
                            : null
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref recipients);
                    OutlookInteropRunner.ReleaseComObject(ref probe);
                    OutlookInteropRunner.ReleaseComObject(ref probeItem);
                }
            },
            ex => new CalendarFreeBusyResult
            {
                Success = false,
                ErrorMessage = $"Failed to read Outlook free/busy information: {ex.Message}"
            });
    }

    /// <summary>
    /// Splits a semicolon- or comma-separated attendee list, dropping blanks.
    /// </summary>
    private static string[] SplitAttendees(string? attendees) =>
        string.IsNullOrWhiteSpace(attendees)
            ? []
            : attendees
                .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToArray();

    /// <summary>
    /// Turns Outlook's slot string into merged non-free intervals. Adjacent slots with the same
    /// status become one period, because a caller looking for a gap wants stretches, not slots.
    /// </summary>
    private static List<FreeBusyPeriodInfo> DecodeBusyPeriods(
        string? availability,
        DateTimeOffset windowStart,
        int intervalMinutes)
    {
        var periods = new List<FreeBusyPeriodInfo>();

        if (string.IsNullOrEmpty(availability))
        {
            return periods;
        }

        int runStart = -1;
        char runStatus = '0';

        for (int index = 0; index <= availability.Length; index++)
        {
            char slot = index < availability.Length ? availability[index] : '0';

            if (slot == runStatus)
            {
                continue;
            }

            if (runStart >= 0 && runStatus != '0')
            {
                periods.Add(new FreeBusyPeriodInfo
                {
                    Start = windowStart.AddMinutes((double)runStart * intervalMinutes),
                    End = windowStart.AddMinutes((double)index * intervalMinutes),
                    Status = DescribeBusyStatus(runStatus)
                });
            }

            runStart = index;
            runStatus = slot;
        }

        return periods;
    }

    private static string DescribeBusyStatus(char slot) => slot switch
    {
        '0' => "free",
        '1' => "tentative",
        '2' => "busy",
        '3' => "outOfOffice",
        '4' => "workingElsewhere",
        _ => "unknown"
    };

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
                BusyStatus = SafeGetInt(() => (int)appointment.BusyStatus),
                IsMeeting = SafeGetBool(() => appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting),
                Attendees = ReadAttendees(appointment, resolveFirst: false),
                IsRecurring = SafeGetBool(() => appointment.IsRecurring),
                RecurrenceState = DescribeRecurrenceStateOf(appointment),
                Recurrence = ReadRecurrence(appointment)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    /// <summary>
    /// Reads the invitee list off an appointment.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <paramref name="resolveFirst"/> asks Outlook to resolve the names against the address book.
    /// That is only wanted when attendees have just been assigned from caller-supplied text; on an
    /// item read back from the store the recipients are already resolved, and resolving again would
    /// be a needless round trip that can prompt.
    /// </para>
    /// <para>
    /// <c>Recipients.ResolveAll</c> returns false when *any* recipient failed, so it cannot say which
    /// one. Each <c>Recipient.Resolved</c> flag is read individually instead - naming the attendee
    /// that failed is the whole point.
    /// </para>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static List<MeetingAttendeeInfo> ReadAttendees(Outlook.AppointmentItem appointment, bool resolveFirst)
    {
        var attendees = new List<MeetingAttendeeInfo>();
        Outlook.Recipients? recipients = null;

        try
        {
            recipients = appointment.Recipients;

            if (recipients == null)
            {
                return attendees;
            }

            if (resolveFirst)
            {
                _ = recipients.ResolveAll();
            }

            int count = recipients.Count;

            for (int index = 1; index <= count; index++)
            {
                Outlook.Recipient? recipient = null;

                try
                {
                    recipient = recipients[index];

                    attendees.Add(new MeetingAttendeeInfo
                    {
                        Name = SafeGet(() => recipient.Name),
                        Address = SafeGet(() => recipient.Address),
                        Type = DescribeRecipientType(SafeGetInt(() => recipient.Type)),
                        ResponseStatus = DescribeResponseStatus(SafeGetInt(() => (int)recipient.MeetingResponseStatus)),
                        Resolved = SafeGetBool(() => recipient.Resolved)
                    });
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref recipient);
                }
            }

            return attendees;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref recipients);
        }
    }

    private static string DescribeRecipientType(int type) => type switch
    {
        (int)Outlook.OlMeetingRecipientType.olOrganizer => "organizer",
        (int)Outlook.OlMeetingRecipientType.olRequired => "required",
        (int)Outlook.OlMeetingRecipientType.olOptional => "optional",
        (int)Outlook.OlMeetingRecipientType.olResource => "resource",
        _ => "unknown"
    };

    private static string DescribeResponseStatus(int status) => status switch
    {
        (int)Outlook.OlResponseStatus.olResponseNone => "none",
        (int)Outlook.OlResponseStatus.olResponseOrganized => "organizer",
        (int)Outlook.OlResponseStatus.olResponseTentative => "tentative",
        (int)Outlook.OlResponseStatus.olResponseAccepted => "accepted",
        (int)Outlook.OlResponseStatus.olResponseDeclined => "declined",
        (int)Outlook.OlResponseStatus.olResponseNotResponded => "notResponded",
        _ => "unknown"
    };

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
            BusyStatus = SafeGetInt(() => (int)appointment.BusyStatus),
            IsRecurring = SafeGetBool(() => appointment.IsRecurring),
            RecurrenceState = DescribeRecurrenceStateOf(appointment)
        };
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static CalendarMutationResult CreateCalendarMutationResult(
        Outlook.AppointmentItem appointment,
        bool updated,
        bool deleted,
        string message,
        string scope = "series")
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
            Scope = scope,
            Message = message
        };
    }

    /// <summary>
    /// Resolves the single occurrence of a recurring series that falls on <paramref name="rawOccurrenceDate"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Two Outlook behaviours make this worth a dedicated method rather than an inline call.
    /// </para>
    /// <para>
    /// First, <c>RecurrencePattern.GetOccurrence</c> matches on the occurrence's exact start, to the
    /// minute, and throws if it is off by one. A caller asking to cancel "Thursday's stand-up" should
    /// not have to know that the stand-up starts at 09:17, so a value with no time of day takes its
    /// time from the series master. A value that does carry a time is used as given, because a caller
    /// who names one is answering a different question and may be targeting an instance that has
    /// already been moved.
    /// </para>
    /// <para>
    /// Second, the item this returns is <em>not</em> the master. Saving it turns that instance into an
    /// exception; deleting it removes only that instance. That is the whole point: an occurrence
    /// carries the master's entry id, so without this the caller's "move one meeting" moved all of
    /// them and reported success.
    /// </para>
    /// <para>
    /// The <c>RecurrencePattern</c> is released here rather than handed back, per Microsoft's guidance
    /// not to hold one across other calls.
    /// </para>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool TryResolveOccurrence(
        Outlook.AppointmentItem master,
        string rawOccurrenceDate,
        out Outlook.AppointmentItem? occurrence,
        out string? errorMessage)
    {
        occurrence = null;
        errorMessage = null;

        if (!DateTimeOffset.TryParse(rawOccurrenceDate, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTimeOffset parsed))
        {
            errorMessage = "occurrenceDate must be a valid ISO date or date/time value.";
            return false;
        }

        if (!SafeGetBool(() => master.IsRecurring))
        {
            errorMessage =
                "occurrenceDate was given but this appointment is not a recurring series, so it has no occurrences. "
                + "Omit occurrenceDate to change the appointment itself.";
            return false;
        }

        DateTime target = parsed.LocalDateTime;

        // A bare date carries no time, so take the series' own time of day rather than midnight -
        // GetOccurrence would otherwise reject every such request.
        bool callerNamedATime = rawOccurrenceDate.Contains(':', StringComparison.Ordinal);
        if (!callerNamedATime)
        {
            DateTimeOffset? masterStart = SafeGetDateTimeOffset(() => master.Start);
            if (masterStart.HasValue)
            {
                target = target.Date + masterStart.Value.LocalDateTime.TimeOfDay;
            }
        }

        Outlook.RecurrencePattern? pattern = null;
        try
        {
            pattern = master.GetRecurrencePattern();
            occurrence = pattern.GetOccurrence(target);
            return true;
        }
        catch (Exception ex)
        {
            // Outlook throws for a date the series does not fall on, and for an instance that has
            // already been cancelled. Both are the caller's question being answered "no", not a
            // failure of the tool - but neither may be turned into a series-wide edit.
            errorMessage =
                $"No occurrence of this series starts at {target:yyyy-MM-dd HH:mm}. "
                + "Check the date against a listing of the series, and note that a cancelled instance no longer exists. "
                + $"Outlook reported: {ex.Message}";
            return false;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref pattern);
        }
    }

    /// <summary>
    /// Asks Outlook for a view of the folder in which each occurrence of a recurring series is a
    /// separate item, restricted to the requested window.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The order matters and is Outlook's, not ours: sort by <c>[Start]</c>, then set
    /// <c>IncludeRecurrences</c>, then <c>Restrict</c>. Setting <c>IncludeRecurrences</c> on an
    /// unsorted collection silently returns masters only, which is the bug this method exists to fix
    /// wearing a disguise.
    /// </para>
    /// <para>
    /// The restriction is an overlap test rather than containment, so a meeting that begins before
    /// the window and runs into it is still returned. Over-inclusive is the only safe direction:
    /// Restrict runs inside Outlook, so anything it wrongly drops reaches the caller as a confident
    /// "nothing is scheduled".
    /// </para>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Items? TryExpandRecurrences(
        Outlook.Items items,
        DateTimeOffset rangeStart,
        DateTimeOffset rangeEnd)
    {
        try
        {
            items.IncludeRecurrences = true;

            string filter = string.Format(
                CultureInfo.InvariantCulture,
                "[Start] <= '{0}' AND [End] >= '{1}'",
                FormatRestrictDate(rangeEnd),
                FormatRestrictDate(rangeStart));

            return items.Restrict(filter);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Formats a bound for Outlook's Jet restriction syntax.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>The user's culture, not the invariant one.</b> Outlook parses a Jet date literal using the
    /// machine's regional settings, so a US-formatted literal is not merely unidiomatic on a European
    /// machine - it is silently misread. Verified on an en-DE machine: <c>09/03/2026 08:41 PM</c>
    /// matched 2 appointments where the equivalent <c>03/09/2026 20:41</c> matched 12. The day and
    /// month swap, the window lands somewhere else entirely, and ten real appointments vanish.
    /// </para>
    /// <para>
    /// It fails quietly and asymmetrically, which is what makes it dangerous: a whole-day window
    /// still works, because midnight survives a mangled time, so the bug only shows up on the
    /// intraday questions - "am I free at 3?" - where a wrong answer matters most.
    /// </para>
    /// <para>
    /// Note this is the opposite of the DASL <c>@SQL=</c> filters used for mail, which take a
    /// culture-independent UTC literal. The two query languages are not interchangeable.
    /// </para>
    /// </remarks>
    private static string FormatRestrictDate(DateTimeOffset value)
        => value.LocalDateTime.ToString("g", CultureInfo.CurrentCulture);

    /// <summary>
    /// Walks an expanded collection with <c>GetFirst</c>/<c>GetNext</c>.
    /// </summary>
    /// <remarks>
    /// Indexing is not an option here: with <c>IncludeRecurrences</c> set, <c>Count</c> is not
    /// meaningful and an endless series has no last item. The scan cap is a guard against exactly
    /// that, not a paging limit.
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void CollectExpandedAppointments(
        Outlook.Items expanded,
        DateTimeOffset? rangeStart,
        DateTimeOffset? rangeEnd,
        int boundedMaxCount,
        bool includeBodyPreview,
        CalendarListResult result)
    {
        const int scanCeiling = 2000;

        object? rawItem = expanded.GetFirst();
        int scanned = 0;

        while (rawItem != null && scanned < scanCeiling && result.Appointments.Count < boundedMaxCount)
        {
            Outlook.AppointmentItem? appointment = null;

            try
            {
                scanned++;
                appointment = rawItem as Outlook.AppointmentItem;

                if (appointment != null && MatchesRange(appointment, rangeStart, rangeEnd))
                {
                    result.Appointments.Add(CreateCalendarSummary(appointment, includeBodyPreview));
                }
            }
            finally
            {
                OutlookInteropRunner.ReleaseComObject(ref appointment);
                OutlookInteropRunner.ReleaseComObject(ref rawItem);
            }

            rawItem = expanded.GetNext();
        }

        if (rawItem != null)
        {
            OutlookInteropRunner.ReleaseComObject(ref rawItem);
        }
    }

    private static bool TryParseRecurrenceType(
        string? recurrenceType,
        out Outlook.OlRecurrenceType parsed)
    {
        switch (recurrenceType?.Trim().ToLowerInvariant())
        {
            case "daily":
                parsed = Outlook.OlRecurrenceType.olRecursDaily;
                return true;
            case "weekly":
                parsed = Outlook.OlRecurrenceType.olRecursWeekly;
                return true;
            case "monthly":
                parsed = Outlook.OlRecurrenceType.olRecursMonthly;
                return true;
            case "yearly":
                parsed = Outlook.OlRecurrenceType.olRecursYearly;
                return true;
            default:
                parsed = Outlook.OlRecurrenceType.olRecursDaily;
                return false;
        }
    }

    private static bool TryParseDayOfWeek(string token, out DayOfWeek day)
        => Enum.TryParse(token.Trim(), ignoreCase: true, out day);

    private static Outlook.OlDaysOfWeek ToDayOfWeekMask(IEnumerable<DayOfWeek> days)
    {
        Outlook.OlDaysOfWeek mask = 0;

        foreach (DayOfWeek day in days)
        {
            mask |= day switch
            {
                DayOfWeek.Sunday => Outlook.OlDaysOfWeek.olSunday,
                DayOfWeek.Monday => Outlook.OlDaysOfWeek.olMonday,
                DayOfWeek.Tuesday => Outlook.OlDaysOfWeek.olTuesday,
                DayOfWeek.Wednesday => Outlook.OlDaysOfWeek.olWednesday,
                DayOfWeek.Thursday => Outlook.OlDaysOfWeek.olThursday,
                DayOfWeek.Friday => Outlook.OlDaysOfWeek.olFriday,
                _ => Outlook.OlDaysOfWeek.olSaturday
            };
        }

        return mask;
    }

    private static List<string> FromDayOfWeekMask(Outlook.OlDaysOfWeek mask)
    {
        var days = new List<string>();

        // Monday first: a weekly pattern read back is nearly always described to a person, and
        // "monday, thursday" reads as a working week where "thursday, monday" does not.
        (Outlook.OlDaysOfWeek Flag, string Name)[] order =
        [
            (Outlook.OlDaysOfWeek.olMonday, "monday"),
            (Outlook.OlDaysOfWeek.olTuesday, "tuesday"),
            (Outlook.OlDaysOfWeek.olWednesday, "wednesday"),
            (Outlook.OlDaysOfWeek.olThursday, "thursday"),
            (Outlook.OlDaysOfWeek.olFriday, "friday"),
            (Outlook.OlDaysOfWeek.olSaturday, "saturday"),
            (Outlook.OlDaysOfWeek.olSunday, "sunday")
        ];

        foreach ((Outlook.OlDaysOfWeek flag, string name) in order)
        {
            if ((mask & flag) == flag)
            {
                days.Add(name);
            }
        }

        return days;
    }

    private static string DescribeRecurrenceType(Outlook.OlRecurrenceType type) => type switch
    {
        Outlook.OlRecurrenceType.olRecursDaily => "daily",
        Outlook.OlRecurrenceType.olRecursWeekly => "weekly",
        Outlook.OlRecurrenceType.olRecursMonthly => "monthly",
        Outlook.OlRecurrenceType.olRecursMonthNth => "monthNth",
        Outlook.OlRecurrenceType.olRecursYearly => "yearly",
        Outlook.OlRecurrenceType.olRecursYearNth => "yearNth",
        _ => "unknown"
    };

    private static string DescribeRecurrenceState(Outlook.OlRecurrenceState state) => state switch
    {
        Outlook.OlRecurrenceState.olApptNotRecurring => "notRecurring",
        Outlook.OlRecurrenceState.olApptMaster => "master",
        Outlook.OlRecurrenceState.olApptOccurrence => "occurrence",
        Outlook.OlRecurrenceState.olApptException => "exception",
        _ => "unknown"
    };

    /// <summary>
    /// Reads the recurrence pattern off an appointment, or null when it is not a series.
    /// </summary>
    /// <remarks>
    /// The pattern is fetched here and released before returning, per Microsoft's documented guidance
    /// for <c>GetRecurrencePattern</c>: a pattern reference held across other work on the parent item
    /// is a known source of corruption. Nothing outside this method ever sees the COM object.
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static RecurrencePatternInfo? ReadRecurrence(Outlook.AppointmentItem appointment)
    {
        if (!SafeGetBool(() => appointment.IsRecurring))
        {
            return null;
        }

        Outlook.RecurrencePattern? pattern = null;
        Outlook.Exceptions? exceptions = null;

        try
        {
            pattern = appointment.GetRecurrencePattern();

            if (pattern == null)
            {
                return null;
            }

            bool noEndDate = SafeGetBool(() => pattern.NoEndDate);
            int occurrences = SafeGetInt(() => pattern.Occurrences);
            Outlook.OlRecurrenceType type = pattern.RecurrenceType;

            exceptions = SafeGetExceptions(pattern);

            return new RecurrencePatternInfo
            {
                RecurrenceType = DescribeRecurrenceType(type),
                Interval = SafeGetInt(() => pattern.Interval),
                DaysOfWeek = type is Outlook.OlRecurrenceType.olRecursWeekly
                    or Outlook.OlRecurrenceType.olRecursMonthNth
                    or Outlook.OlRecurrenceType.olRecursYearNth
                    ? FromDayOfWeekMask(pattern.DayOfWeekMask)
                    : [],
                DayOfMonth = type is Outlook.OlRecurrenceType.olRecursMonthly
                    or Outlook.OlRecurrenceType.olRecursYearly
                    ? SafeGetInt(() => pattern.DayOfMonth)
                    : null,
                MonthOfYear = type is Outlook.OlRecurrenceType.olRecursYearly
                    or Outlook.OlRecurrenceType.olRecursYearNth
                    ? SafeGetInt(() => pattern.MonthOfYear)
                    : null,
                PatternStartDate = SafeGetDateTimeOffset(() => pattern.PatternStartDate),
                // Outlook stores a sentinel end date on an endless series. Reporting it would read as
                // a real end date to anybody who did not also check NoEndDate.
                PatternEndDate = noEndDate ? null : SafeGetDateTimeOffset(() => pattern.PatternEndDate),
                NoEndDate = noEndDate,
                Occurrences = noEndDate ? null : occurrences,
                DurationMinutes = SafeGetInt(() => pattern.Duration),
                ExceptionCount = exceptions == null ? 0 : SafeGetInt(() => exceptions.Count)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref exceptions);
            OutlookInteropRunner.ReleaseComObject(ref pattern);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Exceptions? SafeGetExceptions(Outlook.RecurrencePattern pattern)
    {
        try
        {
            return pattern.Exceptions;
        }
        catch
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string DescribeRecurrenceStateOf(Outlook.AppointmentItem appointment)
    {
        try
        {
            return DescribeRecurrenceState(appointment.RecurrenceState);
        }
        catch
        {
            return "unknown";
        }
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
