using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Reminder discovery (#15).
///
/// <para>
/// Outlook keeps one live collection of everything it intends to remind the user about, across
/// appointments, tasks and flagged mail. Reconstructing that by scanning folders for
/// <c>ReminderSet</c> gets a different and worse answer: it misses flagged mail, it misses the
/// individual instances of a recurring appointment, and it costs far more.
/// </para>
///
/// <para>
/// Three properties of that collection were measured against a real mailbox before any of this was
/// written, and all three are traps that a plausible-looking implementation falls straight into.
/// They are what these tests exist to pin down. Read-only throughout - nothing here dismisses,
/// snoozes or creates a reminder.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailReminder")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookReminderTests(ITestOutputHelper output)
{
    /// <summary>
    /// The baseline: reminders are enumerated, and each arrives with the caption the user would see
    /// and the time it is set for.
    /// </summary>
    [SkippableFact]
    public void ListReminders_ReturnsRemindersWithCaptionAndDueTime()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListReminders();

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.TotalCount == 0, "This profile has no reminders set.");

        foreach (var reminder in result.Reminders)
        {
            Assert.False(string.IsNullOrWhiteSpace(reminder.Caption), "A reminder arrived without a caption.");
            Assert.NotEqual(default, reminder.ReminderTime);
        }

        output.WriteLine($"{result.TotalCount} reminder(s) total, {result.Reminders.Count} returned.");
    }

    /// <summary>
    /// The first trap. <c>Reminder.NextReminderDate</c> reads as the obvious "when will this fire",
    /// but Outlook only populates it once a reminder has actually been snoozed or recurred. On the
    /// test mailbox 152 of 605 reminders return the OLE zero date, 30 December 1899, instead.
    ///
    /// <para>
    /// So the naive implementation returns a complete, confident listing in which a quarter of the
    /// rows are dated 1899. It does not fail and it does not warn. The due time must come from
    /// <c>OriginalReminderDate</c>, which is populated on all of them.
    /// </para>
    ///
    /// <para>
    /// This runs with <c>upcomingOnly: false</c> deliberately. A 1899 row is by definition overdue,
    /// so the default filter would quietly discard exactly the rows this test exists to catch and
    /// the test would pass while the bug shipped - which is what happened when it was first written.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void ListReminders_DoesNotReportTheOleZeroDate()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListReminders(maxCount: 5000, upcomingOnly: false);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Reminders.Count == 0, "This profile has no reminders set.");

        foreach (var reminder in result.Reminders)
        {
            Assert.True(
                reminder.ReminderTime.Year > 1900,
                $"Reminder '{reminder.Caption}' reported {reminder.ReminderTime:yyyy-MM-dd}, which is the OLE zero date rather than a real due time.");
        }

        // NextReminderTime is the snooze time, and is absent unless the reminder is genuinely
        // snoozed. Reporting the zero date here would be the same bug wearing a different name.
        foreach (var reminder in result.Reminders)
        {
            if (reminder.NextReminderTime.HasValue)
            {
                Assert.True(
                    reminder.NextReminderTime.Value.Year > 1900,
                    $"Reminder '{reminder.Caption}' reported a snooze time of {reminder.NextReminderTime:yyyy-MM-dd}.");
            }
        }
    }

    /// <summary>
    /// The second trap. The collection does not arrive in date order - verified against the test
    /// mailbox, where it is emphatically not sorted. Combined with a default limit, taking the
    /// first N gives the caller an arbitrary handful out of the middle of six years of reminders and
    /// presents it as though it meant something.
    /// </summary>
    [SkippableFact]
    public void ListReminders_ReturnsRemindersInChronologicalOrder()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListReminders(maxCount: 100);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Reminders.Count < 2, "Not enough reminders on this profile to establish an ordering.");

        var times = result.Reminders.Select(r => r.ReminderTime).ToList();

        Assert.Equal(times.OrderBy(t => t).ToList(), times);
    }

    /// <summary>
    /// The third trap, and the one with real consequences. Most reminders on a long-lived mailbox
    /// are overdue - 416 of 605 on the test profile, the oldest from 2020. Sorting ascending and
    /// taking the first 50 therefore hands back five-year-old debris in answer to "what am I being
    /// reminded about", while the reminders that are actually coming up never appear.
    ///
    /// <para>
    /// So upcoming reminders are the default, and the count that was left out is reported rather
    /// than quietly dropped. A caller that is shown 50 rows and not told 416 were excluded has been
    /// misled just as surely as by a wrong date.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void ListReminders_DefaultsToUpcomingAndSaysHowManyAreOverdue()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var upcoming = commands.ListReminders();

        Assert.True(upcoming.Success, upcoming.ErrorMessage);
        Skip.If(upcoming.TotalCount == 0, "This profile has no reminders set.");

        foreach (var reminder in upcoming.Reminders)
        {
            Assert.False(
                reminder.IsOverdue,
                $"Reminder '{reminder.Caption}' is overdue but was returned by the upcoming-only default.");
        }

        Assert.Equal(upcoming.TotalCount, upcoming.UpcomingCount + upcoming.OverdueCount);

        Skip.If(upcoming.OverdueCount == 0, "This profile has no overdue reminders.");

        // The overdue ones must be reachable, and must genuinely be the ones held back.
        var all = commands.ListReminders(maxCount: 5000, upcomingOnly: false);
        Assert.True(all.Success, all.ErrorMessage);

        Assert.Equal(upcoming.TotalCount, all.Reminders.Count);
        Assert.Contains(all.Reminders, r => r.IsOverdue);

        output.WriteLine($"{upcoming.OverdueCount} overdue, {upcoming.UpcomingCount} upcoming.");
    }

    /// <summary>
    /// <c>IsVisible</c> is the fourth trap in waiting. It means "the reminder dialog is on screen
    /// right now", not "this reminder is pending" - it was false for all 605 reminders on the test
    /// mailbox. Anything that presented it as pending-ness would report a mailbox stacked with
    /// reminders as having none at all, so the due time and the overdue flag carry that meaning
    /// instead.
    /// </summary>
    [SkippableFact]
    public void ListReminders_DerivesOverdueFromTheDueTimeRatherThanVisibility()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListReminders(maxCount: 5000, upcomingOnly: false);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Reminders.Count == 0, "This profile has no reminders set.");

        var now = DateTime.Now;

        foreach (var reminder in result.Reminders)
        {
            Assert.Equal(reminder.ReminderTime < now, reminder.IsOverdue);
        }
    }

    /// <summary>
    /// maxCount has to bound the returned rows without distorting the counts: the caller needs to
    /// know how much it did not see, or a truncated page reads as the whole picture.
    /// </summary>
    [SkippableFact]
    public void ListReminders_LimitsRowsWithoutMisreportingTheTotal()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var full = commands.ListReminders(maxCount: 5000, upcomingOnly: false);

        Assert.True(full.Success, full.ErrorMessage);
        Skip.If(full.Reminders.Count < 3, "Not enough reminders on this profile to exercise a limit.");

        var limited = commands.ListReminders(maxCount: 2, upcomingOnly: false);

        Assert.True(limited.Success, limited.ErrorMessage);
        Assert.Equal(2, limited.Reminders.Count);
        Assert.Equal(full.TotalCount, limited.TotalCount);

        // Truncation must take the earliest, not an arbitrary two.
        Assert.Equal(
            full.Reminders.Take(2).Select(r => r.Caption),
            limited.Reminders.Select(r => r.Caption));
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
