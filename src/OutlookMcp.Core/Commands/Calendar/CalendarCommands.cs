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
        bool display = false,
        string? requiredAttendees = null,
        string? optionalAttendees = null,
        bool sendInvitation = false)
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
                Attendees = ReadAttendees(appointment, resolveFirst: false)
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
