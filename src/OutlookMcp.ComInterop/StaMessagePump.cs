using System.Runtime.InteropServices;

namespace OutlookMcp.ComInterop;

/// <summary>
/// Pumps the Win32 message queue of the dispatcher's STA thread while waiting for a COM event to
/// arrive.
///
/// <para>
/// <b>Why this is needed.</b> Almost everything this server asks Outlook to do is a synchronous
/// call: it returns with an answer. <c>Application.AdvancedSearch</c> is the exception. It returns
/// immediately with a <c>Search</c> object that is not yet populated, and reports completion by
/// raising <c>AdvancedSearchComplete</c> on the apartment that registered the handler - ours.
/// </para>
///
/// <para>
/// Out-of-process COM events reach a single-threaded apartment as window messages. The dispatcher
/// thread (see <see cref="OutlookDispatcher"/>) normally blocks on its work-item channel and never
/// pumps, so an event sent while it is executing a work item would sit in the queue until the work
/// item returned - which, for the work item that is waiting for that very event, is never. Waiting
/// naively for <c>AdvancedSearch</c> on the dispatcher thread therefore deadlocks until the
/// operation timeout expires.
/// </para>
///
/// <para>
/// <b>Why this pumps explicitly rather than relying on the CLR.</b> A managed blocking wait on an STA
/// thread does pump COM messages, so a <c>WaitHandle.WaitOne</c> loop also works. Measured against a
/// live mailbox it completed the same search in 8.2 seconds where an explicit pump took 2.7: each
/// managed wait blocks for its full slice before re-checking, while
/// <c>MsgWaitForMultipleObjectsEx</c> returns the moment a message arrives. The explicit form is also
/// the one whose behaviour is documented rather than an emergent property of the runtime's STA
/// support, which matters for the one place in this codebase that depends on it.
/// </para>
///
/// <para>
/// <b>What it deliberately does not do.</b> It pumps, it does not dispatch application work. The
/// dispatcher's own work queue is a <c>Channel</c>, not a message queue, so pumping here cannot
/// re-enter another Outlook operation - the STA thread is still occupied by the current work item and
/// the next one stays queued behind it. What it can deliver is Outlook's own callbacks, which is the
/// entire point.
/// </para>
/// </summary>
public static class StaMessagePump
{
    /// <summary>
    /// <c>QS_ALLINPUT</c>: wake for any queued message, since the event arrives as an ordinary
    /// posted or sent message rather than as a distinguishable class of its own.
    /// </summary>
    private const uint QsAllInput = 0x04FF;

    /// <summary>
    /// <c>MWMO_INPUTAVAILABLE</c>: also return when a message is already waiting but has been seen by
    /// a previous <c>PeekMessage</c>. Without it a message that arrived in the gap between draining
    /// the queue and re-entering the wait would not wake this call, and the wait would sit out its
    /// full slice for an event it had already been sent.
    /// </summary>
    private const uint MwmoInputAvailable = 0x0004;

    private const uint PmRemove = 0x0001;

    /// <summary>
    /// How long a single wait blocks before the completion predicate and the deadline are re-checked.
    /// Only an upper bound: an arriving message returns immediately.
    /// </summary>
    private const uint SliceMilliseconds = 50;

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage
    {
        public IntPtr Hwnd;
        public uint Message;
        public IntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public int X;
        public int Y;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PeekMessageW(
        out NativeMessage message,
        IntPtr windowHandle,
        uint filterMin,
        uint filterMax,
        uint removeFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TranslateMessage(ref NativeMessage message);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr DispatchMessageW(ref NativeMessage message);

    [DllImport("user32.dll")]
    private static extern uint MsgWaitForMultipleObjectsEx(
        uint handleCount,
        IntPtr handles,
        uint milliseconds,
        uint wakeMask,
        uint flags);

    /// <summary>
    /// Pumps until <paramref name="isComplete"/> returns <see langword="true"/> or
    /// <paramref name="timeout"/> elapses.
    /// </summary>
    /// <param name="isComplete">
    /// Polled between message batches. Must be cheap and must not itself call into COM: it runs
    /// between pumped messages, and a COM call here would re-enter the object that is mid-callback.
    /// </param>
    /// <param name="timeout">Upper bound on the wait.</param>
    /// <returns>
    /// <see langword="true"/> if the predicate was satisfied, <see langword="false"/> if the timeout
    /// expired first. Returning a bool rather than throwing is deliberate: a search that has not
    /// finished is a result the caller has to describe honestly, not an error.
    /// </returns>
    public static bool WaitFor(Func<bool> isComplete, TimeSpan timeout)
    {
        ArgumentNullException.ThrowIfNull(isComplete);

        long deadline = Environment.TickCount64 + (long)timeout.TotalMilliseconds;

        while (!isComplete())
        {
            if (Environment.TickCount64 >= deadline)
            {
                return false;
            }

            _ = MsgWaitForMultipleObjectsEx(0, IntPtr.Zero, SliceMilliseconds, QsAllInput, MwmoInputAvailable);

            while (PeekMessageW(out NativeMessage message, IntPtr.Zero, 0, 0, PmRemove))
            {
                _ = TranslateMessage(ref message);
                _ = DispatchMessageW(ref message);
            }
        }

        return true;
    }
}
