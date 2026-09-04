using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Tasks;

/// <summary>
/// Outlook task operations.
/// </summary>
[ServiceCategory("task")]
[McpTool("task", Title = "Outlook Task Operations", Destructive = true, Category = "task",
    Description = "Inspect and change Outlook tasks without opening a persistent session. "
    + "Use list to enumerate the default Tasks folder or an explicit Outlook folder path, read to inspect one "
    + "task by entry id or the task currently open or selected in Outlook, create to save a new task, "
    + "update to change named fields on an existing one - fields that are not passed are left alone - and "
    + "delete to remove one. delete needs no confirmation in the ordinary case, because Outlook moves the task "
    + "to Deleted Items where the user can restore it; deleting a task that is already in Deleted Items is "
    + "permanent and requires confirm=true. "
    + "list omits completed tasks unless includeCompleted is true. That default matters: on a real task list "
    + "almost everything is already finished, so a listing that included them would return a page of things "
    + "nobody has to do. completedItemCount reports how many were filtered out, so 'no open tasks' is never "
    + "confused with 'this listing dropped rows'. "
    + "Dates are absent when Outlook has no value for them. Outlook itself stores 1 January 4501 for an unset "
    + "task date; that sentinel is never returned, so a missing dueDate means there is no due date rather than "
    + "a due date in the year 4501. "
    + "status is one of not-started, in-progress, complete, waiting or deferred. Setting status to complete is "
    + "how a task is marked done: Outlook then sets percentComplete to 100 and stamps dateCompleted itself. "
    + "Every task carries an entryId; subjects are not unique and some tasks have no subject at all, so entryId "
    + "is the only reliable handle.")]
public interface ITaskCommands
{
    [ServiceAction("list", Destructive = false)]
    TaskListResult List(
        string? folder = null,
        int maxCount = 25,
        bool includeCompleted = false,
        bool includeBodyPreview = false);

    [ServiceAction("read", Destructive = false)]
    TaskItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveTask = true);

    [ServiceAction("create", Destructive = true)]
    TaskMutationResult Create(
        string subject,
        string? folder = null,
        string? dueDate = null,
        string? startDate = null,
        string? status = null,
        int? percentComplete = null,
        string? importance = null,
        string? categories = null,
        string? body = null,
        bool display = false);

    /// <summary>
    /// Changes named fields on an existing task. Fields that are not passed are left alone.
    /// </summary>
    /// <param name="useActiveTask">Off by default. A mutating action must not fall back to whatever the user has selected in Outlook: the caller chooses the verb and the selection would silently choose the object, so a mistyped id would edit a different task and report success.</param>
    [ServiceAction("update", Destructive = true)]
    TaskMutationResult Update(
        string? entryId = null,
        string? storeId = null,
        string? subject = null,
        string? dueDate = null,
        string? startDate = null,
        string? status = null,
        int? percentComplete = null,
        string? importance = null,
        string? categories = null,
        string? body = null,
        bool useActiveTask = false);

    /// <summary>
    /// Deletes a task.
    ///
    /// <para>
    /// An ordinary delete moves the task to Deleted Items, where the user can restore it, so it is
    /// not gated. Deleting a task that is already in Deleted Items destroys it and requires
    /// <paramref name="confirm"/>.
    /// </para>
    /// </summary>
    /// <param name="confirm">Required only when the task is already in Deleted Items. An ordinary delete ignores it.</param>
    [ServiceAction("delete", Destructive = true)]
    TaskMutationResult Delete(
        string? entryId = null,
        string? storeId = null,
        bool useActiveTask = false,
        bool confirm = false);
}
