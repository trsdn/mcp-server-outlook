using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Recurring appointments (#33).
///
/// <para>
/// The hole this closes is worse than "a feature is missing". Outlook stores a recurring series as a
/// single master item dated at its first occurrence. Listing a calendar without asking Outlook to
/// expand the series therefore returns *nothing* for a weekly stand-up when you ask about next
/// Tuesday - and returns it as a confident empty list. An agent asked "am I free Tuesday at 10?"
/// would have answered yes. That is this project's characteristic failure mode: a call reporting
/// success without having looked at what was asked about.
/// </para>
///
/// <para>
/// Everything here happens in the caller's own calendar, with GUID-named items deleted in a
/// <c>finally</c>. No attendee is ever named, so nothing can be sent to anybody.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CalendarRecurrence")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookRecurrenceTests(ITestOutputHelper output)
{
    /// <summary>
    /// The load-bearing test. A daily series created once must come back as several separate
    /// occurrences on separate days when a range is listed. If expansion is not happening, this
    /// returns one item or none, and the count assertion fails.
    /// </summary>
    [SkippableFact]
    public void List_ExpandsADailySeriesIntoItsOccurrences()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp recurrence test {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(2).Date.AddHours(9);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            recurrenceType: "daily",
            recurrenceCount: 5);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);
            Assert.True(created.IsRecurring);

            var listed = commands.List(
                start: start.AddHours(-1).ToString("o"),
                endTime: start.AddDays(4).AddHours(1).ToString("o"),
                maxCount: 100);

            Assert.True(listed.Success, listed.ErrorMessage);
            Assert.True(listed.RecurringExpanded);

            var mine = listed.Appointments
                .Where(a => string.Equals(a.Subject, subject, StringComparison.Ordinal))
                .ToList();

            output.WriteLine($"Found {mine.Count} occurrence(s) of '{subject}' in the listed range.");
            foreach (var occurrence in mine)
            {
                output.WriteLine($"  {occurrence.Start:yyyy-MM-dd HH:mm} state={occurrence.RecurrenceState}");
            }

            // Five daily occurrences, of which the range covers all five.
            Assert.True(
                mine.Count >= 3,
                $"Expected the daily series to expand into several occurrences, got {mine.Count}. "
                + "One means the series is being reported as its master only, which is the bug.");

            Assert.Equal(mine.Count, mine.Select(a => a.Start?.Date).Distinct().Count());
            Assert.All(mine, a => Assert.Equal("occurrence", a.RecurrenceState));
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Expansion needs a bounded range - a series with no end date has infinitely many occurrences.
    /// A listing without one must say plainly that recurring occurrences are not included rather than
    /// quietly omitting them, which is the same confidently-empty answer in a different disguise.
    /// </summary>
    [SkippableFact]
    public void List_WithoutARange_SaysRecurringOccurrencesAreNotExpanded()
    {
        EnsureOutlookAvailable();

        var listed = new CalendarCommands().List(maxCount: 5);

        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.False(listed.RecurringExpanded);
        Assert.NotNull(listed.Message);
        Assert.Contains("recurring", listed.Message!, StringComparison.OrdinalIgnoreCase);

        output.WriteLine($"Unbounded listing reported: {listed.Message}");
    }

    /// <summary>
    /// Reading the master back must describe the pattern. A caller that cannot see "every day, five
    /// times" cannot tell a series from a single appointment, and will describe one as the other.
    /// </summary>
    [SkippableFact]
    public void Read_ReportsTheRecurrencePattern()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp recurrence read {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(9).Date.AddHours(14);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddHours(1).ToString("o"),
            recurrenceType: "weekly",
            recurrenceInterval: 2,
            recurrenceDaysOfWeek: "monday;thursday",
            recurrenceCount: 6);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var read = commands.Read(entryId: created.EntryId, storeId: created.StoreId, useActiveAppointment: false);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.True(read.IsRecurring);
            Assert.Equal("master", read.RecurrenceState);
            Assert.NotNull(read.Recurrence);
            Assert.Equal("weekly", read.Recurrence!.RecurrenceType);
            Assert.Equal(2, read.Recurrence.Interval);
            Assert.Equal(6, read.Recurrence.Occurrences);
            Assert.False(read.Recurrence.NoEndDate);
            Assert.Contains("monday", read.Recurrence.DaysOfWeek);
            Assert.Contains("thursday", read.Recurrence.DaysOfWeek);

            output.WriteLine(
                $"Pattern: {read.Recurrence.RecurrenceType} every {read.Recurrence.Interval} on "
                + $"{string.Join(",", read.Recurrence.DaysOfWeek)}, {read.Recurrence.Occurrences} times, "
                + $"ending {read.Recurrence.PatternEndDate:yyyy-MM-dd}");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// A non-recurring appointment must stay non-recurring, and must not grow a phantom pattern.
    /// </summary>
    [SkippableFact]
    public void Read_OnAPlainAppointment_ReportsNoRecurrence()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(11).Date.AddHours(11);

        var created = commands.CreateAppointment(
            subject: $"OutlookMcp plain appointment {Guid.NewGuid():N}",
            start: start.ToString("o"),
            endTime: start.AddMinutes(20).ToString("o"));

        try
        {
            Assert.True(created.Success, created.ErrorMessage);
            Assert.False(created.IsRecurring);

            var read = commands.Read(entryId: created.EntryId, storeId: created.StoreId, useActiveAppointment: false);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.False(read.IsRecurring);
            Assert.Equal("notRecurring", read.RecurrenceState);
            Assert.Null(read.Recurrence);
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// An unrecognised recurrence type must be refused before anything is written. Creating a plain
    /// appointment instead, and reporting success, would leave the caller believing they had made a
    /// series - the same class of confidently wrong answer this surface exists to remove.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_WithAnUnknownRecurrenceType_FailsWithoutCreatingAnything()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(13).Date.AddHours(10);

        var created = commands.CreateAppointment(
            subject: $"OutlookMcp bad recurrence {Guid.NewGuid():N}",
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            recurrenceType: "fortnightly");

        try
        {
            Assert.False(created.Success);
            Assert.Null(created.EntryId);
            Assert.NotNull(created.ErrorMessage);
            Assert.Contains("fortnightly", created.ErrorMessage!, StringComparison.Ordinal);

            output.WriteLine($"Rejected as expected: {created.ErrorMessage}");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// A weekly pattern with no day named is ambiguous, and guessing which day the caller meant is
    /// exactly the kind of confident invention that produces a meeting on the wrong day.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_WeeklyWithoutDays_UsesTheStartDay()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(15).Date.AddHours(16);

        var created = commands.CreateAppointment(
            subject: $"OutlookMcp weekly default {Guid.NewGuid():N}",
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            recurrenceType: "weekly",
            recurrenceCount: 3);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var read = commands.Read(entryId: created.EntryId, storeId: created.StoreId, useActiveAppointment: false);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.NotNull(read.Recurrence);
            Assert.Single(read.Recurrence!.DaysOfWeek);
            Assert.Equal(
                start.DayOfWeek.ToString().ToLowerInvariant(),
                read.Recurrence.DaysOfWeek[0]);
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Regression guard for the Jet date-literal format.
    ///
    /// <para>
    /// Expansion restricts the folder inside Outlook, and Outlook parses the date literal using the
    /// machine's regional settings. A US-formatted literal on a European machine has its day and
    /// month swapped, so the window lands somewhere else and the listing comes back empty - a
    /// confident "you have nothing scheduled" for a slot that is booked. It fails only on intraday
    /// windows, because midnight survives a mangled time, so a whole-day test would not catch it.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void List_OverAnIntradayWindow_FindsAnAppointmentInsideIt()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp intraday window {Guid.NewGuid():N}";

        // 13:37 deliberately: a time whose day-of-month and month cannot be confused with each other
        // would hide the bug on half the days of the month.
        DateTimeOffset start = DateTimeOffset.Now.AddDays(4).Date.AddHours(13).AddMinutes(37);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(20).ToString("o"));

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var listed = commands.List(
                start: start.AddMinutes(-15).ToString("o"),
                endTime: start.AddMinutes(35).ToString("o"),
                maxCount: 50);

            Assert.True(listed.Success, listed.ErrorMessage);
            Assert.True(listed.RecurringExpanded);
            Assert.Contains(listed.Appointments, a => a.EntryId == created.EntryId);

            output.WriteLine($"Intraday window returned {listed.Appointments.Count} appointment(s), including the created one.");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Removes a calendar item created by a test. Failures are reported rather than swallowed: an
    /// item left in the owner's real calendar is a defect in the test, not an acceptable outcome.
    /// </summary>
    private void DeleteCreatedItem(CalendarCommands commands, string? entryId, string? storeId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            return;
        }

        var deleted = commands.DeleteAppointment(entryId: entryId, storeId: storeId, useActiveAppointment: false);

        if (!deleted.Success)
        {
            output.WriteLine($"WARNING: could not delete test calendar item {entryId}: {deleted.ErrorMessage}");
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping calendar recurrence test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInterop.NameSpace? session = null;

        try
        {
            session = application!.GetNamespace("MAPI");
            _ = session.Folders.Count;
        }
        catch (Exception ex)
        {
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            if (session is not null && System.Runtime.InteropServices.Marshal.IsComObject(session))
            {
                _ = System.Runtime.InteropServices.Marshal.FinalReleaseComObject(session);
            }
        }
    }
}
