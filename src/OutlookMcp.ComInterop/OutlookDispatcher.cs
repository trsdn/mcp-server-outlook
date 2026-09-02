using System.Threading.Channels;

namespace OutlookMcp.ComInterop;

/// <summary>
/// Process-wide single-threaded apartment (STA) dispatcher that serializes all Outlook COM
/// access behind one long-lived STA thread. Implements #20: previously
/// <c>OutlookInteropRunner.Execute</c> spawned a brand-new STA thread per call, registering and
/// revoking <see cref="OleMessageFilter"/> on every call, which allowed overlapping operations to
/// race on the shared process-wide Outlook RCW and to revoke a filter another in-flight operation
/// was still relying on for RPC_E_SERVERCALL_RETRYLATER retries.
/// </summary>
/// <remarks>
/// <para>
/// This type is deliberately generic (no Outlook COM types appear in its signature) — the
/// Outlook-specific work (resolving <c>Outlook.Application</c>/<c>Outlook.NameSpace</c>,
/// releasing them without final-releasing the shared <c>Application</c> RCW per #19) stays in
/// <c>OutlookMcp.Core.Commands.OutlookInterop.OutlookInteropRunner</c>, which now dispatches its
/// per-call work onto this shared queue instead of owning STA thread lifecycle itself.
/// </para>
/// <para>
/// Uses a <c>Channel&lt;Func&lt;Task&gt;&gt;</c> work queue, <b>bounded</b>
/// to give overlapping callers an explicit back-pressure story: once the queue is full, new
/// callers block (subject to the same overall timeout) instead of piling up unboundedly, and
/// <see cref="OleMessageFilter"/> is
/// registered exactly once for the dispatcher's entire lifetime rather than per operation.
/// </para>
/// </remarks>
public sealed class OutlookDispatcher : IDisposable
{
    /// <summary>
    /// Maximum number of queued-but-not-yet-executing operations before new callers experience
    /// back-pressure (their <see cref="Execute{T}"/> call blocks, subject to the caller-supplied timeout).
    /// Outlook operations are simple/fast MAPI calls, so a modest bound is enough to smooth out
    /// bursts without masking genuine overload.
    /// </summary>
    private const int QueueCapacity = 32;

    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private int _disposed;

    /// <summary>
    /// The operation currently occupying the STA thread, or <see langword="null"/> when idle.
    /// Written only by the dispatcher thread, read by arbitrary callers, hence
    /// <see cref="Volatile"/> access. Used to detect a wedged dispatcher (see
    /// <see cref="Execute{T}"/>).
    /// </summary>
    private InFlightOperation? _inFlight;

    /// <summary>
    /// A work item occupying the STA thread, together with the wall-clock instant at which its
    /// own caller stops waiting for it.
    /// </summary>
    private sealed record InFlightOperation(string Name, DateTime DeadlineUtc);

    /// <summary>
    /// The single process-wide dispatcher instance. Outlook itself is a single shared,
    /// already-running <c>Application</c> per user session (see ADR-002), so there is exactly
    /// one dispatcher for the whole process — not one per operation, and not one per caller.
    /// </summary>
    public static OutlookDispatcher Shared { get; } = new();

    private OutlookDispatcher()
    {
        _workQueue = Channel.CreateBounded<Func<Task>>(new BoundedChannelOptions(QueueCapacity)
        {
            SingleReader = true,
            SingleWriter = false,
            FullMode = BoundedChannelFullMode.Wait
        });
        _shutdownCts = new CancellationTokenSource();

        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        _staThread = new Thread(() => RunMessagePump(started))
        {
            IsBackground = true,
            Name = "OutlookDispatcher"
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        // Wait for the filter to be registered before returning, so the first caller's work item
        // is guaranteed to be queued only after the dispatcher is actually ready to process it.
        started.Task.GetAwaiter().GetResult();
    }

    /// <summary>
    /// Shuts down the dispatcher's STA thread. Since <see cref="Shared"/> is a process-wide
    /// singleton intended to live for the process lifetime, this is provided mainly for test
    /// scenarios and orderly process shutdown, not for routine per-call use.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
        {
            return;
        }

        _workQueue.Writer.TryComplete();
        _shutdownCts.Cancel();
        _ = _staThread.Join(TimeSpan.FromSeconds(10));
        _shutdownCts.Dispose();
    }

    /// <summary>
    /// Dispatches <paramref name="operation"/> onto the single shared STA thread and blocks until
    /// it completes or <paramref name="timeout"/> elapses. The timeout covers both queuing
    /// (back-pressure while the bounded queue is full) and execution, so callers get one
    /// predictable deadline regardless of which phase is slow.
    /// </summary>
    /// <typeparam name="T">The operation's result type.</typeparam>
    /// <param name="operationName">A short, human-readable name used in timeout messages.</param>
    /// <param name="operation">The COM work to run on the dispatcher's STA thread.</param>
    /// <param name="timeout">The maximum time to wait for the operation to be queued and executed.</param>
    /// <exception cref="TimeoutException">
    /// Thrown if the operation could not be queued or did not complete within <paramref name="timeout"/>,
    /// or immediately if the dispatcher is currently blocked by an earlier operation that has already
    /// exceeded its own deadline and cannot be cancelled.
    /// </exception>
    /// <exception cref="ObjectDisposedException">Thrown if the dispatcher has been shut down.</exception>
    public T Execute<T>(string operationName, Func<T> operation, TimeSpan timeout)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(OutlookDispatcher));

        // Quarantine a wedged dispatcher (#19). A blocking cross-apartment COM call cannot be
        // interrupted: Thread.Abort is unsupported on .NET 5+, and OleMessageFilter only governs
        // *incoming* calls, so it cannot rescue an outgoing call parked on a modal Outlook Object
        // Model Guard dialog. Once the in-flight operation is past its own deadline its caller has
        // already thrown, nobody can ever observe its result, and it will hold this thread until a
        // person dismisses the dialog. Queueing behind it would make every subsequent caller burn
        // its own full timeout -- DefaultOperationTimeout is five minutes -- before reporting a
        // generic failure that names the wrong operation. Fail immediately and name what is stuck.
        // This is transient: the check clears itself as soon as the stuck operation returns.
        InFlightOperation? stuck = Volatile.Read(ref _inFlight);
        if (stuck is not null && DateTime.UtcNow > stuck.DeadlineUtc)
        {
            throw new TimeoutException(
                $"{operationName} was rejected because the Outlook dispatcher is blocked by '{stuck.Name}', " +
                $"which exceeded its own deadline at {stuck.DeadlineUtc:O} and cannot be cancelled. " +
                "This usually means Outlook is showing a modal dialog (for example an Object Model Guard " +
                "security prompt) that a person must dismiss, or that Outlook is busy with a long " +
                "synchronization. Check Outlook for a visible dialog; operations resume automatically " +
                "once the blocked call returns.");
        }

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var timeoutCts = new CancellationTokenSource(timeout);
        DateTime deadlineUtc = DateTime.UtcNow + timeout;

        try
        {
            var writeTask = _workQueue.Writer.WriteAsync(() =>
            {
                // Do not run work whose caller has already given up (#19). Outlook operations are
                // not all read-only -- mail.send, mail.delete and folder.create mutate the user's
                // mailbox -- so executing an abandoned item performs a side effect that the caller
                // was told had timed out and may legitimately have retried. The deadline is checked
                // here rather than via timeoutCts because that token source is disposed as soon as
                // Execute returns, which for a timed-out caller is before this item is reached.
                // This narrows the window rather than closing it: a caller can still time out in
                // the instant between this check and the call below. That residual race is
                // unavoidable without cancellable COM, and is orders of magnitude smaller than
                // the queue wait it replaces.
                if (DateTime.UtcNow >= deadlineUtc)
                {
                    tcs.TrySetCanceled();
                    return Task.CompletedTask;
                }

                Volatile.Write(ref _inFlight, new InFlightOperation(operationName, deadlineUtc));
                try
                {
                    tcs.TrySetResult(operation());
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
                finally
                {
                    Volatile.Write(ref _inFlight, null);
                }

                return Task.CompletedTask;
            }, timeoutCts.Token);

            if (writeTask.IsCompleted)
            {
                writeTask.GetAwaiter().GetResult();
            }
            else
            {
                writeTask.AsTask().GetAwaiter().GetResult();
            }
        }
        catch (ChannelClosedException)
        {
            throw new ObjectDisposedException(nameof(OutlookDispatcher), "The Outlook dispatcher has been shut down.");
        }
        catch (OperationCanceledException)
        {
            throw new TimeoutException(
                $"{operationName} timed out after {timeout.TotalSeconds} seconds waiting for a free slot on the Outlook dispatcher queue " +
                "(too many overlapping Outlook operations in flight).");
        }

        try
        {
            return tcs.Task.WaitAsync(timeoutCts.Token).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException)
        {
            throw new TimeoutException($"{operationName} timed out after {timeout.TotalSeconds} seconds while running on the Outlook dispatcher.");
        }
    }

    private void RunMessagePump(TaskCompletionSource started)
    {
        try
        {
            // Registered exactly once for the entire lifetime of this dispatcher thread — the
            // core fix for #20 (previously registered/revoked on every single Execute call).
            OleMessageFilter.Register();
        }
        catch (Exception ex)
        {
            started.TrySetException(ex);
            return;
        }

        started.TrySetResult();

        try
        {
            while (true)
            {
                try
                {
                    if (!_workQueue.Reader.WaitToReadAsync(_shutdownCts.Token).AsTask().GetAwaiter().GetResult())
                    {
                        // Writer completed (process shutdown) — exit gracefully.
                        break;
                    }

                    while (_workQueue.Reader.TryRead(out var work))
                    {
                        try
                        {
                            work().GetAwaiter().GetResult();
                        }
                        catch (Exception)
                        {
                            // Individual work items report failure via their own TaskCompletionSource;
                            // keep pumping so later queued operations are not starved by one failure.
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    // Drain any remaining queued work so in-flight callers get a prompt result
                    // instead of waiting out their own timeout after shutdown was requested.
                    while (_workQueue.Reader.TryRead(out var remainingWork))
                    {
                        try
                        {
                            remainingWork().GetAwaiter().GetResult();
                        }
                        catch (Exception)
                        {
                            // Already captured by the work item's own TaskCompletionSource.
                        }
                    }

                    break;
                }
            }
        }
        finally
        {
            OleMessageFilter.Revoke();
        }
    }
}
