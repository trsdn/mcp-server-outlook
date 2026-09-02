using Xunit;

namespace OutlookMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for OleMessageFilter registration and revocation.
/// Tests verify that the message filter can be registered/revoked without errors.
///
/// NOTE: These tests verify the registration mechanism but don't test actual
/// COM retry behavior (that requires Outlook and would be integration tests).
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class OleMessageFilterTests
{
    [Fact]
    public void Register_OnStaThread_DoesNotThrow()
    {
        // Arrange & Act & Assert
        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Filter registration failed: {ex.Message}", ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    [Fact]
    public void RegisterAndRevoke_MultipleTimes_DoesNotThrow()
    {
        // Arrange & Act & Assert
        var thread = new Thread(() =>
        {
            // First registration
            OleMessageFilter.Register();
            OleMessageFilter.Revoke();

            // Second registration (simulates reuse)
            OleMessageFilter.Register();
            OleMessageFilter.Revoke();
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    [Fact]
    public void Revoke_WithoutRegister_DoesNotThrow()
    {
        // Revoke without prior Register should not crash
        // Arrange & Act & Assert - Should handle gracefully
        var thread = new Thread(OleMessageFilter.Revoke);

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    /// <summary>
    /// REGRESSION TEST: MessagePending MUST return PENDINGMSG_WAITDEFPROCESS, NOT
    /// PENDINGMSG_WAITNOPROCESS.
    ///
    /// Per the Win32 PENDINGMSG enumeration (objidl.h):
    ///   PENDINGMSG_CANCELCALL    = 0  cancel the outgoing call
    ///   PENDINGMSG_WAITNOPROCESS = 1  wait for the return and do NOT dispatch the message
    ///   PENDINGMSG_WAITDEFPROCESS = 2 wait and dispatch the message
    /// https://learn.microsoft.com/windows/win32/api/objidl/ne-objidl-pendingmsg
    ///
    /// Returning WAITNOPROCESS (1) blocks all inbound COM message processing while an outgoing
    /// call is in progress. If the callee fires a re-entrant callback during that call, the
    /// callback is queued but never dispatched: the callee waits for the callback, the STA thread
    /// waits for the callee, and the two deadlock.
    ///
    /// Returning WAITDEFPROCESS (2) lets COM dispatch the pending inbound call to
    /// HandleInComingCall, which decides whether to accept it (SERVERCALL_ISHANDLED during normal
    /// operations) or reject it with SERVERCALL_RETRYLATER during a long operation so the caller's
    /// RetryRejectedCall backoff runs.
    /// </summary>
    [Fact]
    public void MessagePending_ReturnValue_MustBe_WaitDefProcess()
    {
        // The IOleMessageFilter interface is internal, so we verify the constant value via
        // reflection on the compiled method body — simpler: we verify by checking the
        // registered filter doesn't use the blocking value (2).
        //
        // We instantiate the filter and call MessagePending via the interface.
        // Use reflection to access the internal interface implementation.
        const int PENDINGMSG_WAITNOPROCESS = 1;
        const int PENDINGMSG_WAITDEFPROCESS = 2;

        var returnValue = -1;
        Exception? threadException = null;

        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();

                // The filter implements IOleMessageFilter which is internal.
                // We can verify via the public static IsRegistered and the logical behavior:
                // After Register(), the filter IS the active message filter for this thread.
                //
                // Verify that the filter is registered (prerequisite for the bug to manifest).
                Assert.True(OleMessageFilter.IsRegistered, "Filter must be registered to have any effect");

                // Use reflection to invoke MessagePending on the filter instance.
                // The filter class is internal, but we can get to it via the assembly.
                var filterType = typeof(OleMessageFilter);
                var iOleMsgFilterType = filterType.Assembly.GetType(
                    "OutlookMcp.ComInterop.IOleMessageFilter");
                Assert.NotNull(iOleMsgFilterType);

                // Create a filter instance and call MessagePending
                var filterInstance = Activator.CreateInstance(filterType);
                Assert.NotNull(filterInstance);
                var method = iOleMsgFilterType.GetMethod("MessagePending");
                Assert.NotNull(method);

                returnValue = (int)method.Invoke(filterInstance, [IntPtr.Zero, 1000, 1])!;
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                threadException = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (threadException != null) throw new InvalidOperationException($"Thread exception: {threadException.Message}", threadException);

        // REGRESSION: if this returns WAITNOPROCESS (1), an outgoing call that provokes a
        // re-entrant inbound callback deadlocks - the callback is queued but never dispatched,
        // so the callee waits on us while we wait on the callee.
        Assert.NotEqual(PENDINGMSG_WAITNOPROCESS, returnValue);
        Assert.Equal(PENDINGMSG_WAITDEFPROCESS, returnValue);
    }
}





