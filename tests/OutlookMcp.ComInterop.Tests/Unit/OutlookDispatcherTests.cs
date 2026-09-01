using System.Collections.Concurrent;
using OutlookMcp.ComInterop.Session;
using Xunit;

namespace OutlookMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for <see cref="OutlookDispatcher"/>'s serialization behavior. These tests exercise
/// only the dispatcher's STA-thread/queue mechanics with plain delegates — no Outlook COM
/// automation is involved — so they fall under Rule 30's documented exception for "pure
/// algorithmic utilities with zero COM dependency". They are the regression coverage called for
/// by #20's acceptance criteria: "Test issuing overlapping Outlook operations and asserting no
/// InvalidComObjectException" (modeled here as "no torn/interleaved shared state", since without
/// real Outlook installed there is no COM object to invalidate).
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "OutlookDispatcher")]
public class OutlookDispatcherTests
{
    [Fact]
    public async Task Execute_OverlappingCallers_RunsSerializedOnSingleThread()
    {
        // Regression test for #20: OutlookInteropRunner used to spawn a brand-new STA thread per
        // call. If two operations ever ran concurrently against shared state, updates would
        // interleave. OutlookDispatcher.Shared serializes all work onto one STA thread, so
        // concurrently-issued operations must still observe strictly serialized execution:
        // no two operations' "critical sections" overlap, and every one of them runs on the
        // same managed thread ID.
        const int callerCount = 20;
        var observedThreadIds = new ConcurrentBag<int>();
        var concurrentEntries = 0;
        var maxObservedConcurrency = 0;
        var sync = new object();

        var callers = Enumerable.Range(0, callerCount).Select(i => Task.Run(() =>
            OutlookDispatcher.Shared.Execute($"overlap-test-{i}", () =>
            {
                int current = Interlocked.Increment(ref concurrentEntries);
                lock (sync)
                {
                    maxObservedConcurrency = Math.Max(maxObservedConcurrency, current);
                }

                observedThreadIds.Add(Environment.CurrentManagedThreadId);

                // Simulate a small amount of "COM work" so overlapping callers would actually
                // have a window to race in if serialization were broken.
                Thread.Sleep(5);

                Interlocked.Decrement(ref concurrentEntries);
                return i;
            }, TimeSpan.FromSeconds(30))))
            .ToArray();

        await Task.WhenAll(callers);

        Assert.Equal(1, maxObservedConcurrency);
        Assert.Single(observedThreadIds.Distinct());

        var results = callers.Select(t => t.Result).OrderBy(x => x).ToArray();
        Assert.Equal(Enumerable.Range(0, callerCount), results);
    }

    [Fact]
    public void Execute_WhenOperationThrows_PropagatesExceptionWithoutStoppingDispatcher()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            OutlookDispatcher.Shared.Execute<int>("throwing-op", () => throw new InvalidOperationException("boom"), TimeSpan.FromSeconds(30)));

        Assert.Equal("boom", ex.Message);

        // Dispatcher must still be usable after a failed operation — one bad work item must not
        // wedge the shared STA thread for subsequent callers.
        int result = OutlookDispatcher.Shared.Execute("after-throw", () => 42, TimeSpan.FromSeconds(30));
        Assert.Equal(42, result);
    }

    [Fact]
    public void Execute_WhenTimeoutElapsesDuringExecution_ThrowsTimeoutException()
    {
        Assert.Throws<TimeoutException>(() =>
            OutlookDispatcher.Shared.Execute("slow-op", () =>
            {
                Thread.Sleep(500);
                return 1;
            }, TimeSpan.FromMilliseconds(50)));

        // The dispatcher's single STA thread is still busy running the slow work item's
        // Thread.Sleep(500) when this timeout fires (Execute's own timeout only bounds the
        // *caller's* wait, not the in-flight work item). Give it a moment to finish before
        // asserting the dispatcher is usable again, so this test isn't racing the prior
        // work item's completion.
        Thread.Sleep(600);

        int result = OutlookDispatcher.Shared.Execute("after-timeout", () => 7, TimeSpan.FromSeconds(30));
        Assert.Equal(7, result);
    }
}
