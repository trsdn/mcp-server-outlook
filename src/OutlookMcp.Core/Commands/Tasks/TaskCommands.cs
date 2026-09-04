using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Tasks;

/// <summary>
/// Outlook task operations (#14).
///
/// <para>
/// Two facts about real task folders drive everything here, and both were measured before any of
/// it was written. First, <b>Outlook does not use null for "no date"</b>: an unset
/// <c>DueDate</c>, <c>StartDate</c> or <c>DateCompleted</c> reads as 1 January 4501. On the mailbox
/// this was built against that is 260 of 274 due dates, so reporting it verbatim would date 95% of
/// a listing to the 46th century. Second, <b>nearly every task is finished</b> - 271 of those 274 -
/// so a listing that does not exclude completed tasks by default returns a page of things nobody
/// has to do, with the handful that matter buried in it.
/// </para>
///
/// <para>
/// Both are the same failure: an answer that is confidently wrong rather than visibly empty. The
/// filtered-out tasks are therefore counted in <see cref="TaskListResult.CompletedItemCount"/>
/// rather than silently dropped, so "no open tasks" can never be confused with "this listing lost
/// rows".
/// </para>
/// </summary>
public class TaskCommands : ITaskCommands
{
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> FolderAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["tasks"] = Outlook.OlDefaultFolders.olFolderTasks,
            ["task"] = Outlook.OlDefaultFolders.olFolderTasks,
            ["todo"] = Outlook.OlDefaultFolders.olFolderTasks,
            ["current"] = Outlook.OlDefaultFolders.olFolderTasks
        };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public TaskListResult List(
        string? folder = null,
        int maxCount = 25,
        bool includeCompleted = false,
        bool includeBodyPreview = false)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 1000);

        return OutlookInteropRunner.Execute(
            "OutlookTaskList",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? taskFolder = null;
                Outlook.Items? items = null;

                try
                {
                    taskFolder = ResolveTaskFolder(application, session, folder, ref explorer);
                    if (taskFolder == null)
                    {
                        return new TaskListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = taskFolder.Items;
                    int totalItemCount = SafeGetInt(() => items.Count);
                    TrySortByDueDate(items);

                    var result = new TaskListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => taskFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(taskFolder),
                        TotalItemCount = totalItemCount,
                        IncludedCompleted = includeCompleted
                    };

                    int scanned = 0;

                    for (int index = 1; index <= totalItemCount; index++)
                    {
                        if (result.Tasks.Count >= boundedMaxCount)
                        {
                            break;
                        }

                        object? rawItem = null;

                        try
                        {
                            rawItem = items[index];
                            scanned++;

                            if (rawItem is not Outlook.TaskItem task)
                            {
                                // A Tasks folder can hold task requests and anything else a user
                                // filed there. Counted rather than ignored so the totals add up.
                                result.SkippedItemCount++;
                                continue;
                            }

                            var summary = CreateTaskSummary(task, includeBodyPreview);

                            if (summary.Complete && !includeCompleted)
                            {
                                result.CompletedItemCount++;
                                result.SkippedItemCount++;
                                continue;
                            }

                            if (summary.Complete)
                            {
                                result.CompletedItemCount++;
                            }

                            result.Tasks.Add(summary);
                        }
                        catch (COMException)
                        {
                            // The item exists but could not be read - a corrupt row, or one the
                            // Object Model Guard refuses. Counted, because the alternative is a
                            // listing that is quietly short.
                            scanned++;
                            result.SkippedItemCount++;
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ScannedItemCount = scanned;
                    result.ReturnedCount = result.Tasks.Count;
                    result.Truncated = scanned < totalItemCount;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref taskFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new TaskListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook tasks: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public TaskItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveTask = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookTaskRead",
            (application, session) =>
            {
                var resolved = ResolveTaskItem(application, session, entryId, storeId, useActiveTask);

                try
                {
                    if (resolved.Task == null)
                    {
                        return new TaskItemResult
                        {
                            Success = true,
                            HasItem = false
                        };
                    }

                    return CreateTaskItemResult(resolved.Task);
                }
                finally
                {
                    resolved.Release();
                }
            },
            ex => new TaskItemResult
            {
                Success = false,
                HasItem = false,
                ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                    ? $"Failed to inspect the active Outlook task: {ex.Message}"
                    : $"Failed to inspect the requested Outlook task: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public TaskMutationResult Create(
        string subject,
        string? folder = null,
        string? dueDate = null,
        string? startDate = null,
        string? status = null,
        int? percentComplete = null,
        string? importance = null,
        string? categories = null,
        string? body = null,
        bool display = false)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            return new TaskMutationResult
            {
                Success = false,
                ErrorMessage = "A subject is required to create a task."
            };
        }

        // Everything that can be rejected is rejected before Outlook is touched, so a bad value
        // cannot leave a half-written task behind.
        if (!TryParseFields(
                dueDate,
                startDate,
                status,
                percentComplete,
                importance,
                out DateTime? parsedDueDate,
                out DateTime? parsedStartDate,
                out Outlook.OlTaskStatus? parsedStatus,
                out Outlook.OlImportance? parsedImportance,
                out string? validationError))
        {
            return new TaskMutationResult
            {
                Success = false,
                ErrorMessage = validationError
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookTaskCreate",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? taskFolder = null;
                Outlook.Items? items = null;
                object? createdItem = null;
                Outlook.TaskItem? task = null;

                try
                {
                    taskFolder = ResolveTaskFolder(application, session, folder, ref explorer);
                    if (taskFolder == null)
                    {
                        return new TaskMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = taskFolder.Items;
                    createdItem = items.Add(Outlook.OlItemType.olTaskItem);
                    task = createdItem as Outlook.TaskItem;

                    if (task == null)
                    {
                        return new TaskMutationResult
                        {
                            Success = false,
                            ErrorMessage = "Outlook did not return a task item for the new task."
                        };
                    }

                    task.Subject = subject;

                    ApplyTaskUpdates(
                        task,
                        parsedDueDate,
                        parsedStartDate,
                        parsedStatus,
                        percentComplete,
                        parsedImportance,
                        categories,
                        body);

                    task.Save();

                    if (display)
                    {
                        task.Display(false);
                    }

                    var result = CreateTaskMutationResult(task, "Saved Outlook task.");
                    result.Saved = true;
                    result.Displayed = display;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref task);
                    OutlookInteropRunner.ReleaseComObject(ref createdItem);
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref taskFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new TaskMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create the Outlook task: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public TaskMutationResult Update(
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
        bool useActiveTask = true)
    {
        bool hasUpdates =
            subject != null ||
            dueDate != null ||
            startDate != null ||
            status != null ||
            percentComplete != null ||
            importance != null ||
            categories != null ||
            body != null;

        if (!hasUpdates)
        {
            return new TaskMutationResult
            {
                Success = false,
                ErrorMessage = "At least one task field must be provided for update."
            };
        }

        if (!TryParseFields(
                dueDate,
                startDate,
                status,
                percentComplete,
                importance,
                out DateTime? parsedDueDate,
                out DateTime? parsedStartDate,
                out Outlook.OlTaskStatus? parsedStatus,
                out Outlook.OlImportance? parsedImportance,
                out string? validationError))
        {
            return new TaskMutationResult
            {
                Success = false,
                ErrorMessage = validationError
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookTaskUpdate",
            (application, session) =>
            {
                var resolved = ResolveTaskItem(application, session, entryId, storeId, useActiveTask);

                try
                {
                    if (resolved.Task == null)
                    {
                        return new TaskMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedMessage(entryId, "update")
                        };
                    }

                    if (subject != null)
                    {
                        resolved.Task.Subject = subject;
                    }

                    ApplyTaskUpdates(
                        resolved.Task,
                        parsedDueDate,
                        parsedStartDate,
                        parsedStatus,
                        percentComplete,
                        parsedImportance,
                        categories,
                        body);

                    resolved.Task.Save();

                    var result = CreateTaskMutationResult(resolved.Task, "Updated Outlook task.");
                    result.Saved = true;
                    result.Updated = true;
                    return result;
                }
                finally
                {
                    resolved.Release();
                }
            },
            ex => new TaskMutationResult
            {
                Success = false,
                ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                    ? $"Failed to update the active Outlook task: {ex.Message}"
                    : $"Failed to update the requested Outlook task: {ex.Message}"
            });
    }

    /// <summary>
    /// Deletes a task.
    ///
    /// <para>
    /// <c>useActiveTask</c> defaults to false here, unlike read and update. Falling back to
    /// whatever happens to be selected in Outlook is a convenience when reading and a hazard when
    /// deleting: a delete call with a mistyped id would otherwise remove a different task entirely
    /// and report success.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public TaskMutationResult Delete(
        string? entryId = null,
        string? storeId = null,
        bool useActiveTask = false)
    {
        if (string.IsNullOrWhiteSpace(entryId) && !useActiveTask)
        {
            return new TaskMutationResult
            {
                Success = false,
                ErrorMessage = "An entryId is required to delete a task. "
                    + "Pass useActiveTask: true to delete the task currently open or selected in Outlook."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookTaskDelete",
            (application, session) =>
            {
                var resolved = ResolveTaskItem(application, session, entryId, storeId, useActiveTask);

                try
                {
                    if (resolved.Task == null)
                    {
                        return new TaskMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedMessage(entryId, "delete")
                        };
                    }

                    // Read the identifying fields before the delete, because afterwards the item is
                    // gone and every property on it throws.
                    var result = CreateTaskMutationResult(resolved.Task, "Deleted Outlook task.");

                    resolved.Task.Delete();
                    result.Deleted = true;
                    return result;
                }
                finally
                {
                    resolved.Release();
                }
            },
            ex => new TaskMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to delete the Outlook task: {ex.Message}"
            });
    }

    /// <summary>
    /// A resolved task together with the COM objects that had to be held to reach it. Returning
    /// them as one value keeps the release list in one place: the caller cannot forget the explorer
    /// it never asked for.
    /// </summary>
    private sealed class ResolvedTask
    {
        public Outlook.TaskItem? Task { get; set; }

        public Outlook.Inspector? Inspector { get; set; }

        public Outlook.Explorer? Explorer { get; set; }

        public Outlook.Selection? Selection { get; set; }

        public object? RawItem { get; set; }

        public void Release()
        {
            Outlook.TaskItem? task = Task;
            Outlook.Inspector? inspector = Inspector;
            Outlook.Explorer? explorer = Explorer;
            Outlook.Selection? selection = Selection;
            object? rawItem = RawItem;

            OutlookInteropRunner.ReleaseComObject(ref task);
            OutlookInteropRunner.ReleaseComObject(ref rawItem);
            OutlookInteropRunner.ReleaseComObject(ref selection);
            OutlookInteropRunner.ReleaseComObject(ref explorer);
            OutlookInteropRunner.ReleaseComObject(ref inspector);

            Task = null;
            Inspector = null;
            Explorer = null;
            Selection = null;
            RawItem = null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ResolvedTask ResolveTaskItem(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? entryId,
        string? storeId,
        bool useActiveTask)
    {
        var resolved = new ResolvedTask();

        if (!string.IsNullOrWhiteSpace(entryId))
        {
            try
            {
                resolved.RawItem = session.GetItemFromID(
                    entryId,
                    string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
                resolved.Task = resolved.RawItem as Outlook.TaskItem;
            }
            catch (COMException)
            {
                // An id that does not resolve. Reported by the caller as a refusal rather than
                // silently falling through to whatever is selected in the UI.
            }

            return resolved;
        }

        if (!useActiveTask)
        {
            return resolved;
        }

        resolved.Inspector = application.ActiveInspector();
        if (resolved.Inspector != null)
        {
            resolved.RawItem = resolved.Inspector.CurrentItem;
            if (resolved.RawItem is Outlook.TaskItem openTask)
            {
                resolved.Task = openTask;
                return resolved;
            }

            object? notATask = resolved.RawItem;
            OutlookInteropRunner.ReleaseComObject(ref notATask);
            resolved.RawItem = null;
        }

        resolved.Explorer = application.ActiveExplorer();
        if (resolved.Explorer != null)
        {
            resolved.Selection = resolved.Explorer.Selection;
            if (resolved.Selection != null && resolved.Selection.Count > 0)
            {
                resolved.RawItem = resolved.Selection[1];
                if (resolved.RawItem is Outlook.TaskItem selectedTask)
                {
                    resolved.Task = selectedTask;
                }
            }
        }

        return resolved;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static TaskItemResult CreateTaskItemResult(Outlook.TaskItem task)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = task.Parent as Outlook.MAPIFolder;

            return new TaskItemResult
            {
                Success = true,
                HasItem = true,
                EntryId = SafeGet(() => task.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                Subject = BuildSubject(SafeGet(() => task.Subject)),
                Status = DescribeStatus(SafeGetStatus(task)),
                PercentComplete = SafeGetInt(() => task.PercentComplete),
                Complete = SafeGetBool(() => task.Complete),
                DueDate = SafeGetTaskDate(() => task.DueDate),
                StartDate = SafeGetTaskDate(() => task.StartDate),
                DateCompleted = SafeGetTaskDate(() => task.DateCompleted),
                Importance = DescribeImportance(SafeGetImportance(task)),
                Owner = NullIfBlank(SafeGet(() => task.Owner)),
                Categories = NullIfBlank(SafeGet(() => task.Categories)),
                ReminderSet = SafeGetBool(() => task.ReminderSet),
                ReminderTime = SafeGetBool(() => task.ReminderSet)
                    ? SafeGetTaskDate(() => task.ReminderTime)
                    : null,
                FolderPath = OutlookInteropRunner.GetFolderPath(parentFolder),
                BodyPreview = NullIfBlank(OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => task.Body)))
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static TaskSummaryInfo CreateTaskSummary(Outlook.TaskItem task, bool includeBodyPreview)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = task.Parent as Outlook.MAPIFolder;

            return new TaskSummaryInfo
            {
                EntryId = SafeGet(() => task.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                Subject = BuildSubject(SafeGet(() => task.Subject)),
                Status = DescribeStatus(SafeGetStatus(task)),
                PercentComplete = SafeGetInt(() => task.PercentComplete),
                Complete = SafeGetBool(() => task.Complete),
                DueDate = SafeGetTaskDate(() => task.DueDate),
                StartDate = SafeGetTaskDate(() => task.StartDate),
                DateCompleted = SafeGetTaskDate(() => task.DateCompleted),
                Importance = DescribeImportance(SafeGetImportance(task)),
                Owner = NullIfBlank(SafeGet(() => task.Owner)),
                Categories = NullIfBlank(SafeGet(() => task.Categories)),
                ReminderSet = SafeGetBool(() => task.ReminderSet),
                ReminderTime = SafeGetBool(() => task.ReminderSet)
                    ? SafeGetTaskDate(() => task.ReminderTime)
                    : null,
                BodyPreview = includeBodyPreview
                    ? NullIfBlank(OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => task.Body)))
                    : null
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static TaskMutationResult CreateTaskMutationResult(Outlook.TaskItem task, string message)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = task.Parent as Outlook.MAPIFolder;

            return new TaskMutationResult
            {
                Success = true,
                Message = message,
                EntryId = SafeGet(() => task.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                Subject = BuildSubject(SafeGet(() => task.Subject)),
                Status = DescribeStatus(SafeGetStatus(task)),
                PercentComplete = SafeGetInt(() => task.PercentComplete),
                DueDate = SafeGetTaskDate(() => task.DueDate),
                FolderPath = OutlookInteropRunner.GetFolderPath(parentFolder)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    private static void ApplyTaskUpdates(
        Outlook.TaskItem task,
        DateTime? dueDate,
        DateTime? startDate,
        Outlook.OlTaskStatus? status,
        int? percentComplete,
        Outlook.OlImportance? importance,
        string? categories,
        string? body)
    {
        // Only fields that were actually passed are written. A parameter left unset must leave the
        // stored value alone, otherwise changing a status would clear a due date.
        //
        // Order matters: status is written before percentComplete. Outlook derives one from the
        // other - setting status to complete forces percentComplete to 100 - so writing status last
        // would silently overwrite an explicit percentage.
        if (status != null)
        {
            task.Status = status.Value;
        }

        if (percentComplete != null)
        {
            task.PercentComplete = percentComplete.Value;
        }

        // Outlook refuses a start date later than the due date, so when both are supplied the due
        // date goes first.
        if (dueDate != null)
        {
            task.DueDate = dueDate.Value;
        }

        if (startDate != null)
        {
            task.StartDate = startDate.Value;
        }

        if (importance != null)
        {
            task.Importance = importance.Value;
        }

        if (categories != null)
        {
            task.Categories = categories;
        }

        if (body != null)
        {
            task.Body = body;
        }
    }

    /// <summary>
    /// Validates and converts every caller-supplied value up front, before any COM object exists.
    /// A value Outlook would reject has to be refused rather than half-applied: a create that threw
    /// midway would leave a partly written task in the user's task list.
    /// </summary>
    private static bool TryParseFields(
        string? dueDate,
        string? startDate,
        string? status,
        int? percentComplete,
        string? importance,
        out DateTime? parsedDueDate,
        out DateTime? parsedStartDate,
        out Outlook.OlTaskStatus? parsedStatus,
        out Outlook.OlImportance? parsedImportance,
        out string? error)
    {
        parsedStartDate = null;
        parsedStatus = null;
        parsedImportance = null;

        if (!TryParseTaskDate(dueDate, nameof(dueDate), out parsedDueDate, out error))
        {
            return false;
        }

        if (!TryParseTaskDate(startDate, nameof(startDate), out parsedStartDate, out error))
        {
            return false;
        }

        if (status != null)
        {
            if (!TryParseStatus(status, out Outlook.OlTaskStatus statusValue))
            {
                error = $"Unsupported task status '{status}'. "
                    + "Supported values: not-started, in-progress, complete, waiting, deferred.";
                return false;
            }

            parsedStatus = statusValue;
        }

        if (percentComplete is < 0 or > 100)
        {
            error = $"percentComplete must be between 0 and 100; got {percentComplete}.";
            return false;
        }

        if (importance != null)
        {
            if (!TryParseImportance(importance, out Outlook.OlImportance importanceValue))
            {
                error = $"Unsupported importance '{importance}'. Supported values: low, normal, high.";
                return false;
            }

            parsedImportance = importanceValue;
        }

        return true;
    }

    private static bool TryParseTaskDate(
        string? value,
        string parameterName,
        out DateTime? parsed,
        out string? error)
    {
        parsed = null;
        error = null;

        if (value == null)
        {
            return true;
        }

        // An empty string clears the date. Outlook has no null for this, so its own "none" marker
        // is written back instead.
        if (string.IsNullOrWhiteSpace(value))
        {
            parsed = NoTaskDate;
            return true;
        }

        if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime result))
        {
            error = $"Could not parse {parameterName} '{value}'. Use a date such as 2026-03-14 or 2026-03-14T09:00.";
            return false;
        }

        parsed = result;
        return true;
    }

    private static bool TryParseStatus(string status, out Outlook.OlTaskStatus parsed)
    {
        switch (status.Trim().ToLowerInvariant().Replace(" ", "-"))
        {
            case "not-started":
            case "notstarted":
                parsed = Outlook.OlTaskStatus.olTaskNotStarted;
                return true;

            case "in-progress":
            case "inprogress":
                parsed = Outlook.OlTaskStatus.olTaskInProgress;
                return true;

            case "complete":
            case "completed":
            case "done":
                parsed = Outlook.OlTaskStatus.olTaskComplete;
                return true;

            case "waiting":
            case "waiting-on-others":
                parsed = Outlook.OlTaskStatus.olTaskWaiting;
                return true;

            case "deferred":
                parsed = Outlook.OlTaskStatus.olTaskDeferred;
                return true;

            default:
                parsed = Outlook.OlTaskStatus.olTaskNotStarted;
                return false;
        }
    }

    private static bool TryParseImportance(string importance, out Outlook.OlImportance parsed)
    {
        switch (importance.Trim().ToLowerInvariant())
        {
            case "low":
                parsed = Outlook.OlImportance.olImportanceLow;
                return true;

            case "normal":
            case "medium":
                parsed = Outlook.OlImportance.olImportanceNormal;
                return true;

            case "high":
                parsed = Outlook.OlImportance.olImportanceHigh;
                return true;

            default:
                parsed = Outlook.OlImportance.olImportanceNormal;
                return false;
        }
    }

    private static string DescribeStatus(Outlook.OlTaskStatus? status) => status switch
    {
        Outlook.OlTaskStatus.olTaskNotStarted => "not-started",
        Outlook.OlTaskStatus.olTaskInProgress => "in-progress",
        Outlook.OlTaskStatus.olTaskComplete => "complete",
        Outlook.OlTaskStatus.olTaskWaiting => "waiting",
        Outlook.OlTaskStatus.olTaskDeferred => "deferred",
        _ => "unknown"
    };

    private static string DescribeImportance(Outlook.OlImportance? importance) => importance switch
    {
        Outlook.OlImportance.olImportanceLow => "low",
        Outlook.OlImportance.olImportanceNormal => "normal",
        Outlook.OlImportance.olImportanceHigh => "high",
        _ => "unknown"
    };

    /// <summary>Outlook's stand-in for "this task date was never set".</summary>
    private static readonly DateTime NoTaskDate = new(4501, 1, 1);

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? ResolveTaskFolder(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? folder,
        ref Outlook.Explorer? explorer)
        => OutlookInteropRunner.ResolveFolder(
            application,
            session,
            string.IsNullOrWhiteSpace(folder) ? "tasks" : folder,
            FolderAliases,
            ref explorer);

    private static string BuildUnknownFolderMessage(string? folder)
    {
        const string supportedFolders = "current, tasks, or an Outlook folder path";
        return string.IsNullOrWhiteSpace(folder)
            ? $"Could not resolve the Outlook task folder. Supported folder values: {supportedFolders}."
            : $"Unsupported Outlook task folder '{folder}'. Supported folder values: {supportedFolders}.";
    }

    private static string BuildUnresolvedMessage(string? entryId, string operation)
        => string.IsNullOrWhiteSpace(entryId)
            ? $"Could not resolve an active Outlook task to {operation}. "
                + "Open or select a task in Outlook, or pass an entryId."
            : $"Could not resolve the Outlook task '{entryId}' to {operation}. "
                + "The entry id may be wrong, or the item may live in another store - pass its storeId as well.";

    /// <summary>
    /// A label that is never blank. A task saved with no subject is a real thing in a real task
    /// list, and a row with nothing to show is a row the caller cannot render or disambiguate.
    /// </summary>
    private static string BuildSubject(string? subject)
        => string.IsNullOrWhiteSpace(subject) ? "(task with no subject)" : subject;

    private static string? NullIfBlank(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value;

    private static string? SafeGet(Func<string?> getter)
    {
        try
        {
            return getter();
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static int SafeGetInt(Func<int> getter)
    {
        try
        {
            return getter();
        }
        catch (COMException)
        {
            return 0;
        }
    }

    private static bool SafeGetBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch (COMException)
        {
            return false;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.OlTaskStatus? SafeGetStatus(Outlook.TaskItem task)
    {
        try
        {
            return task.Status;
        }
        catch (COMException)
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.OlImportance? SafeGetImportance(Outlook.TaskItem task)
    {
        try
        {
            return task.Importance;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <summary>
    /// Reads a task date, treating Outlook's two "no date" sentinels as absent.
    ///
    /// <para>
    /// Neither is <c>default(DateTime)</c>. Outlook's own "none" marker for a task date is
    /// 1 January 4501, and an unset OLE date surfaces as 30 December 1899. Both look like perfectly
    /// ordinary dates to a caller, so letting either through produces a listing that is confidently
    /// and uniformly wrong - on the mailbox this was built against, 260 of 274 due dates.
    /// </para>
    /// </summary>
    private static DateTimeOffset? SafeGetTaskDate(Func<DateTime> getter)
    {
        try
        {
            DateTime value = getter();

            if (value.Year >= 4000 || value.Year <= 1900)
            {
                return null;
            }

            return new DateTimeOffset(value);
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <summary>
    /// Sorts by due date so that a limited listing returns the most urgent tasks rather than an
    /// arbitrary handful. Best effort: a store that refuses the sort still yields a usable listing,
    /// just in its native order.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void TrySortByDueDate(Outlook.Items items)
    {
        try
        {
            items.Sort("[DueDate]", false);
        }
        catch (COMException)
        {
            // Some stores refuse to sort on this property. Not worth failing the call over.
        }
    }
}
