using OutlookMcp.Core.Commands.OutlookInterop;
using System.Globalization;
using OutlookMcp.Core.Commands.Tasks;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Task items: list, read, create, update and delete (#14).
///
/// <para>
/// <b>Everything asserted here was measured against a real Tasks folder before it was written.</b>
/// That folder holds 274 items, and two facts about it shape the whole design:
/// </para>
///
/// <para>
/// <b>Outlook does not use null for "no date".</b> 260 of the 274 have a <c>DueDate</c> of
/// 1 January 4501 and 272 have that <c>StartDate</c>. Reported verbatim, 95% of a task listing
/// would be dated to the 46th century - a confidently wrong answer of exactly the kind this project
/// keeps finding. The same sentinel is already normalised for mail flags in <c>MailCommands</c>;
/// tasks are where it actually comes from.
/// </para>
///
/// <para>
/// <b>Nearly every task is finished.</b> 271 of the 274 are complete. A listing that does not
/// exclude completed tasks by default returns a page of things nobody has to do, and the three that
/// matter are invisible. So the default is to omit them - which makes the test for it load-bearing
/// in both directions, because a default that silently discards rows is precisely what defeated the
/// reminders test in #15. Both halves are asserted on the same task.
/// </para>
///
/// <para>
/// <b>Mutation safety.</b> Tasks created here carry a GUID scratch prefix in the subject, live in
/// the default Tasks folder, and are deleted in <c>finally</c> with a prefix sweep afterwards. No
/// pre-existing task is ever updated or deleted: every destructive assertion is made against an
/// item the test created moments earlier.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Task")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookTaskTests(ITestOutputHelper output)
{
    /// <summary>
    /// The baseline: tasks come back, every one carries the id needed to address it, and nothing
    /// scanned goes missing. Subjects are not unique - "Follow up" appears many times - so the id
    /// is the only handle.
    /// </summary>
    [SkippableFact]
    public void List_ReturnsTasksThatCanEachBeAddressedById()
    {
        EnsureOutlookAvailable();

        var result = new TaskCommands().List(maxCount: 1000, includeCompleted: true);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Tasks.Count == 0, "This profile has no tasks to list.");

        foreach (var task in result.Tasks)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(task.EntryId),
                $"Task '{task.Subject}' arrived without an entry id.");

            Assert.False(
                string.IsNullOrWhiteSpace(task.Subject),
                $"The task with entry id '{task.EntryId}' has nothing to display.");
        }

        // Everything the scan looked at is either returned or explicitly accounted for. A listing
        // that quietly drops rows is wrong in the one way a caller cannot detect.
        Assert.Equal(result.ScannedItemCount, result.Tasks.Count + result.SkippedItemCount);

        output.WriteLine(
            $"total={result.TotalItemCount} scanned={result.ScannedItemCount} "
            + $"returned={result.ReturnedCount} skipped={result.SkippedItemCount} "
            + $"truncated={result.Truncated} from {result.FolderPath}");
    }

    /// <summary>
    /// The trap that would have shipped. Outlook stores 1 January 4501 for an unset task date, and
    /// on the mailbox this was built against that is 260 of 274 due dates.
    ///
    /// <para>
    /// <c>includeCompleted</c> is passed explicitly here. Leaving it at its default would let the
    /// listing filter away most of the folder before this assertion ever saw it - the same way an
    /// <c>upcomingOnly</c> default silently defeated the reminders test in #15. A test must not be
    /// able to pass because the code under test hid the rows it was looking for.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void List_NeverReportsOutlooksNoDateSentinelAsARealDate()
    {
        EnsureOutlookAvailable();

        var result = new TaskCommands().List(maxCount: 1000, includeCompleted: true);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Tasks.Count == 0, "This profile has no tasks to list.");

        int withoutDueDate = 0;

        foreach (var task in result.Tasks)
        {
            AssertNotASentinelDate(task.DueDate, nameof(task.DueDate), task.Subject);
            AssertNotASentinelDate(task.StartDate, nameof(task.StartDate), task.Subject);
            AssertNotASentinelDate(task.DateCompleted, nameof(task.DateCompleted), task.Subject);

            if (task.DueDate == null)
            {
                withoutDueDate++;
            }
        }

        output.WriteLine(
            $"{withoutDueDate} of {result.Tasks.Count} task(s) report no due date, "
            + "rather than a date in the year 4501.");
    }

    /// <summary>
    /// Completed tasks are excluded by default and included on request, asserted on one task that
    /// the test moves between the two states itself.
    ///
    /// <para>
    /// This cannot be asserted against pre-existing tasks: doing so would need a folder that happens
    /// to contain both kinds, and a test that only passes on one mailbox is a test that stops
    /// testing. Creating the task and completing it makes the assertion hold on any profile.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void List_OmitsCompletedTasksByDefault_AndReturnsThemWhenAsked()
    {
        EnsureOutlookAvailable();

        var commands = new TaskCommands();
        string subject = ScratchSubject();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = commands.Create(subject: subject);
            Assert.True(created.Success, created.ErrorMessage);
            entryId = created.EntryId;
            storeId = created.StoreId;

            Assert.True(
                ListContains(commands, entryId, includeCompleted: false),
                "A task that is not started is missing from the default listing.");

            var completed = commands.Update(entryId, storeId, status: "complete", useActiveTask: false);
            Assert.True(completed.Success, completed.ErrorMessage);

            // Confirmed from the item itself, not from the update's own return value.
            var afterComplete = commands.Read(entryId, storeId, useActiveTask: false);
            Assert.True(afterComplete.Success, afterComplete.ErrorMessage);
            Assert.True(afterComplete.Complete, "The task did not actually become complete.");
            Assert.Equal(100, afterComplete.PercentComplete);

            Assert.False(
                ListContains(commands, entryId, includeCompleted: false),
                "A completed task is still in the default listing, which is meant to omit them.");

            Assert.True(
                ListContains(commands, entryId, includeCompleted: true),
                "A completed task is missing even from a listing that asked for completed tasks.");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
            SweepScratchTasks(commands);
        }
    }

    /// <summary>
    /// The full mutation lifecycle, verified from the outside at every step by re-reading the item
    /// rather than trusting the return value of the call that made the change.
    /// </summary>
    [SkippableFact]
    public void Create_Update_Delete_ReallyChangeTheTasksFolder()
    {
        EnsureOutlookAvailable();

        var commands = new TaskCommands();
        string subject = ScratchSubject();
        string dueDate = DateTime.Today.AddDays(3).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = commands.Create(
                subject: subject,
                dueDate: dueDate,
                body: "Created by the OutlookMcp integration test.",
                importance: "high");

            Assert.True(created.Success, created.ErrorMessage);
            Assert.False(string.IsNullOrWhiteSpace(created.EntryId), "create returned no entry id.");

            entryId = created.EntryId;
            storeId = created.StoreId;
            output.WriteLine($"Created: {entryId}");

            var afterCreate = commands.Read(entryId, storeId, useActiveTask: false);
            Assert.True(afterCreate.Success, afterCreate.ErrorMessage);
            Assert.True(afterCreate.HasItem);
            Assert.Equal(subject, afterCreate.Subject);
            Assert.NotNull(afterCreate.DueDate);
            Assert.Equal(DateTime.Today.AddDays(3).Date, afterCreate.DueDate!.Value.Date);
            Assert.Equal("high", afterCreate.Importance);

            var updated = commands.Update(
                entryId,
                storeId,
                status: "in-progress",
                percentComplete: 50,
                useActiveTask: false);

            Assert.True(updated.Success, updated.ErrorMessage);

            var afterUpdate = commands.Read(entryId, storeId, useActiveTask: false);
            Assert.True(afterUpdate.Success, afterUpdate.ErrorMessage);
            Assert.Equal("in-progress", afterUpdate.Status);
            Assert.Equal(50, afterUpdate.PercentComplete);

            // Fields that were not passed to update must survive it. An update implemented as
            // "write every parameter" would blank these, and a test that only checked the changed
            // field would not notice.
            Assert.Equal(subject, afterUpdate.Subject);
            Assert.NotNull(afterUpdate.DueDate);
        }
        finally
        {
            if (entryId != null)
            {
                var deleted = commands.Delete(entryId, storeId);
                output.WriteLine($"Delete: success={deleted.Success} {deleted.ErrorMessage}");
                Assert.True(deleted.Success, deleted.ErrorMessage);
            }

            SweepScratchTasks(commands);
        }

        if (entryId != null)
        {
            var afterDelete = commands.Read(entryId, storeId, useActiveTask: false);
            Assert.False(
                afterDelete.Success && afterDelete.HasItem,
                "The task is still readable after delete reported success.");
        }
    }

    /// <summary>
    /// A task created with no due date must report no due date - not the sentinel, and not today.
    /// This is the create-side half of <see cref="List_NeverReportsOutlooksNoDateSentinelAsARealDate"/>,
    /// and it holds on any profile because the test supplies the task.
    /// </summary>
    [SkippableFact]
    public void Create_WithoutADueDate_ReportsNoDueDateRatherThanASentinel()
    {
        EnsureOutlookAvailable();

        var commands = new TaskCommands();
        string subject = ScratchSubject();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = commands.Create(subject: subject);
            Assert.True(created.Success, created.ErrorMessage);
            entryId = created.EntryId;
            storeId = created.StoreId;

            var read = commands.Read(entryId, storeId, useActiveTask: false);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.True(read.HasItem);

            Assert.Null(read.DueDate);
            Assert.Null(read.StartDate);
            Assert.Null(read.DateCompleted);
            Assert.False(read.Complete);
            Assert.Equal("not-started", read.Status);
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
            SweepScratchTasks(commands);
        }
    }

    /// <summary>
    /// Update has to refuse an id it cannot resolve rather than inventing a task. Silently creating
    /// one would be the worst possible outcome of a typo.
    /// </summary>
    [SkippableFact]
    public void Update_RefusesAnIdThatDoesNotResolve()
    {
        EnsureOutlookAvailable();

        var commands = new TaskCommands();
        int before = commands.List(maxCount: 1000, includeCompleted: true).TotalItemCount;

        var result = commands.Update(
            entryId: "0000000000000000000000000000000000000000000000",
            subject: "should never be written",
            useActiveTask: false);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));

        int after = commands.List(maxCount: 1000, includeCompleted: true).TotalItemCount;
        Assert.Equal(before, after);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Delete refuses to fall back to whatever is selected in the Outlook window. That fallback is
    /// a convenience when reading and a hazard when deleting: a mistyped id would otherwise remove
    /// a different task entirely and report success.
    /// </summary>
    [SkippableFact]
    public void Delete_RefusesWithoutAnIdRatherThanUsingTheSelection()
    {
        EnsureOutlookAvailable();

        var result = new TaskCommands().Delete(entryId: null, storeId: null);

        Assert.False(result.Success);
        Assert.False(result.Deleted);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// An unparseable due date is refused, and refused before anything is written. Accepting it and
    /// storing "today", or storing the sentinel, would both be silent corruption.
    /// </summary>
    [SkippableFact]
    public void Create_RefusesAnUnparseableDueDate()
    {
        EnsureOutlookAvailable();

        var commands = new TaskCommands();
        int before = commands.List(maxCount: 1000, includeCompleted: true).TotalItemCount;

        var result = commands.Create(subject: ScratchSubject(), dueDate: "not a date");

        Assert.False(result.Success);
        Assert.False(result.Saved);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));

        int after = commands.List(maxCount: 1000, includeCompleted: true).TotalItemCount;
        Assert.Equal(before, after);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    private static void AssertNotASentinelDate(DateTimeOffset? value, string field, string subject)
    {
        if (value == null)
        {
            return;
        }

        Assert.True(
            value.Value.Year is > 1900 and < 4000,
            $"{field} on task '{subject}' reads as {value.Value:yyyy-MM-dd}, which is one of "
            + "Outlook's 'no date' sentinels rather than a real date. It should be absent.");
    }

    private static bool ListContains(TaskCommands commands, string? entryId, bool includeCompleted)
    {
        var listing = commands.List(maxCount: 1000, includeCompleted: includeCompleted);
        Assert.True(listing.Success, listing.ErrorMessage);
        return listing.Tasks.Any(t => t.EntryId == entryId);
    }

    private void DeleteIfCreated(TaskCommands commands, string? entryId, string? storeId)
    {
        if (entryId == null)
        {
            return;
        }

        var deleted = commands.Delete(entryId, storeId);
        output.WriteLine($"Delete: success={deleted.Success} {deleted.ErrorMessage}");
    }

    /// <summary>
    /// Removes anything this file created that a <c>finally</c> block failed to remove. Reported
    /// loudly rather than swallowed: leftovers here are real rows in the user's task list.
    /// </summary>
    private void SweepScratchTasks(TaskCommands commands)
    {
        var listing = commands.List(maxCount: 1000, includeCompleted: true);
        if (!listing.Success)
        {
            output.WriteLine($"Sweep could not list tasks: {listing.ErrorMessage}");
            return;
        }

        var leftovers = listing.Tasks
            .Where(t => t.Subject.StartsWith(ScratchPrefix, StringComparison.Ordinal))
            .ToList();

        foreach (var task in leftovers)
        {
            var deleted = commands.Delete(task.EntryId, task.StoreId);
            output.WriteLine($"Sweep: {task.Subject} -> success={deleted.Success} {deleted.ErrorMessage}");
        }

        if (leftovers.Count > 0)
        {
            output.WriteLine($"SWEEP removed {leftovers.Count} leftover scratch task(s).");
        }
    }

    private const string ScratchPrefix = "mcp-task-test-";

    private static string ScratchSubject() => $"{ScratchPrefix}{Guid.NewGuid():N}";

    private static void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        // Plain decrement, never FinalReleaseComObject: this is the shared Outlook.Application and
        // final-releasing it breaks every other holder in the process (#19, #116).
        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
