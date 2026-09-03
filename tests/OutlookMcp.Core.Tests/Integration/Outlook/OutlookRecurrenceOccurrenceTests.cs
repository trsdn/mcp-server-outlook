using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Single-occurrence edits and cancellations on a recurring series (#33, slice 2).
///
/// <para>
/// The trap this closes is specific and silent. Every occurrence of a series carries the *master's*
/// entry id, so <c>calendar.update-appointment --entry-id ...</c> against an occurrence moved the
/// entire series while looking exactly like a single-instance edit. Cancelling one stand-up wiped
/// out every stand-up. Nothing reported an error, because as far as Outlook was concerned the caller
/// asked for exactly that.
/// </para>
///
/// <para>
/// So the fix is not only "make single occurrences editable", it is "make the scope of the edit
/// something the caller states rather than something they discover afterwards". Naming an
/// occurrenceDate opts into the single-instance edit; omitting it keeps the series-wide behaviour;
/// naming one on an item that is not a series is refused rather than quietly ignored.
/// </para>
///
/// <para>
/// Everything happens in the caller's own calendar with GUID-named items removed in a
/// <c>finally</c>. No attendee is ever named, so nothing can be sent to anybody.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CalendarRecurrence")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookRecurrenceOccurrenceTests(ITestOutputHelper output)
{
    /// <summary>
    /// The load-bearing test for updating. One occurrence gets a new location; every other
    /// occurrence must keep the original. If the edit is leaking to the series - the bug - every
    /// occurrence carries the new location and the "others unchanged" assertion fails.
    /// </summary>
    [SkippableFact]
    public void UpdateAppointment_WithAnOccurrenceDate_ChangesOnlyThatOccurrence()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp occurrence update {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(2).Date.AddHours(9);
        DateTimeOffset target = start.AddDays(2);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            location: "Original room",
            recurrenceType: "daily",
            recurrenceCount: 5);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);
            Assert.True(created.IsRecurring);

            var updated = commands.UpdateAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                location: "Moved room",
                occurrenceDate: target.ToString("o"));

            Assert.True(updated.Success, updated.ErrorMessage);
            Assert.True(updated.Updated);
            Assert.Equal("occurrence", updated.Scope);

            var listed = commands.List(
                start: start.AddHours(-1).ToString("o"),
                endTime: start.AddDays(5).ToString("o"),
                maxCount: 100);

            Assert.True(listed.Success, listed.ErrorMessage);

            var mine = listed.Appointments
                .Where(a => string.Equals(a.Subject, subject, StringComparison.Ordinal))
                .OrderBy(a => a.Start)
                .ToList();

            foreach (var occurrence in mine)
            {
                output.WriteLine($"  {occurrence.Start:yyyy-MM-dd HH:mm} state={occurrence.RecurrenceState}");
            }

            Assert.True(mine.Count >= 4, $"Expected the series to survive the single-occurrence edit, got {mine.Count}.");

            var moved = mine.Where(a => a.Start?.Date == target.Date).ToList();
            var untouched = mine.Where(a => a.Start?.Date != target.Date).ToList();

            Assert.Single(moved);
            Assert.NotEmpty(untouched);

            // The edited instance became an exception; the rest are still plain occurrences.
            Assert.Equal("exception", moved[0].RecurrenceState);
            Assert.All(untouched, a => Assert.Equal("occurrence", a.RecurrenceState));
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// The load-bearing test for cancelling. Removing one occurrence must leave the rest of the
    /// series standing. Before this existed, deleting by entry id removed all of them.
    /// </summary>
    [SkippableFact]
    public void DeleteAppointment_WithAnOccurrenceDate_RemovesOnlyThatOccurrence()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp occurrence delete {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(2).Date.AddHours(11);
        DateTimeOffset target = start.AddDays(2);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            recurrenceType: "daily",
            recurrenceCount: 5);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var deleted = commands.DeleteAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                occurrenceDate: target.ToString("o"));

            Assert.True(deleted.Success, deleted.ErrorMessage);
            Assert.True(deleted.Deleted);
            Assert.Equal("occurrence", deleted.Scope);

            var listed = commands.List(
                start: start.AddHours(-1).ToString("o"),
                endTime: start.AddDays(5).ToString("o"),
                maxCount: 100);

            Assert.True(listed.Success, listed.ErrorMessage);

            var mine = listed.Appointments
                .Where(a => string.Equals(a.Subject, subject, StringComparison.Ordinal))
                .ToList();

            output.WriteLine($"{mine.Count} occurrence(s) left after cancelling one of five.");

            // The cancelled date is gone...
            Assert.DoesNotContain(mine, a => a.Start?.Date == target.Date);

            // ...and the rest of the series is emphatically still there. Asserting only the absence
            // would pass just as happily if the whole series had been wiped out, which is the bug.
            Assert.True(mine.Count >= 3, $"Expected the rest of the series to survive, got {mine.Count}.");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Naming an occurrence on an item that has no recurrence must be refused. Ignoring the argument
    /// and editing the item anyway would be the worst outcome: the caller asked to touch one instance
    /// of something and would be told, truthfully but misleadingly, that it succeeded.
    /// </summary>
    [SkippableFact]
    public void UpdateAppointment_WithAnOccurrenceDateOnASingleAppointment_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp non-recurring {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(3).Date.AddHours(14);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"));

        try
        {
            Assert.True(created.Success, created.ErrorMessage);
            Assert.False(created.IsRecurring);

            var updated = commands.UpdateAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                location: "Should never be applied",
                occurrenceDate: start.ToString("o"));

            Assert.False(updated.Success);
            Assert.False(updated.Updated);
            Assert.NotNull(updated.ErrorMessage);
            Assert.Contains("recurring", updated.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {updated.ErrorMessage}");

            // And the item really was left alone.
            var read = commands.Read(entryId: created.EntryId, storeId: created.StoreId, useActiveAppointment: false);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.NotEqual("Should never be applied", read.Location);
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// A date the series does not fall on has no occurrence. Outlook throws for this; the caller must
    /// get a clear refusal rather than an interop stack trace, and above all must not have the series
    /// edited as a consolation prize.
    /// </summary>
    [SkippableFact]
    public void UpdateAppointment_WithADateTheSeriesDoesNotFallOn_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp missing occurrence {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(2).Date.AddHours(16);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            recurrenceType: "daily",
            recurrenceCount: 3);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            // Well past the third and final occurrence.
            var updated = commands.UpdateAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                location: "Nowhere",
                occurrenceDate: start.AddDays(40).ToString("o"));

            Assert.False(updated.Success);
            Assert.False(updated.Updated);
            Assert.NotNull(updated.ErrorMessage);
            Assert.Contains("occurrence", updated.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {updated.ErrorMessage}");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Outlook's GetOccurrence wants the occurrence's exact start, to the minute, and throws if the
    /// time is off by so much as one. A caller naming a bare date is asking an entirely reasonable
    /// question - "cancel Thursday's stand-up" - and must not have to know the series' start time to
    /// get an answer. A date with no time takes its time of day from the series.
    /// </summary>
    [SkippableFact]
    public void DeleteAppointment_WithADateOnlyOccurrence_TakesTheTimeFromTheSeries()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp date-only occurrence {Guid.NewGuid():N}";

        // 09:17, a time nobody would guess, so passing a bare date can only work if the time is
        // being read off the series rather than defaulted to midnight.
        DateTimeOffset start = DateTimeOffset.Now.AddDays(2).Date.AddHours(9).AddMinutes(17);
        DateTimeOffset target = start.AddDays(1);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(25).ToString("o"),
            recurrenceType: "daily",
            recurrenceCount: 4);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var deleted = commands.DeleteAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                occurrenceDate: target.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture));

            Assert.True(deleted.Success, deleted.ErrorMessage);
            Assert.Equal("occurrence", deleted.Scope);

            var listed = commands.List(
                start: start.AddHours(-1).ToString("o"),
                endTime: start.AddDays(4).ToString("o"),
                maxCount: 100);

            Assert.True(listed.Success, listed.ErrorMessage);

            var mine = listed.Appointments
                .Where(a => string.Equals(a.Subject, subject, StringComparison.Ordinal))
                .ToList();

            Assert.DoesNotContain(mine, a => a.Start?.Date == target.Date);
            Assert.True(mine.Count >= 2, $"Expected the rest of the series to survive, got {mine.Count}.");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Omitting occurrenceDate must keep working exactly as before - the whole series is the target.
    /// This is the backwards-compatibility guard: the new parameter must not have quietly turned
    /// every existing series edit into a single-instance one.
    /// </summary>
    [SkippableFact]
    public void UpdateAppointment_WithoutAnOccurrenceDate_StillEditsTheWholeSeries()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"OutlookMcp series edit {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.AddDays(2).Date.AddHours(15);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            location: "Original room",
            recurrenceType: "daily",
            recurrenceCount: 4);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var updated = commands.UpdateAppointment(
                entryId: created.EntryId,
                storeId: created.StoreId,
                location: "Series room");

            Assert.True(updated.Success, updated.ErrorMessage);
            Assert.Equal("series", updated.Scope);

            var read = commands.Read(entryId: created.EntryId, storeId: created.StoreId, useActiveAppointment: false);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.Equal("Series room", read.Location);
            Assert.True(read.IsRecurring);
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
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
