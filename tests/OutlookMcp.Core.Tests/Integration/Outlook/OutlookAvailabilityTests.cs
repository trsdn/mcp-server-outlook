using System.Diagnostics;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Reaching a running Outlook regardless of Running Object Table registration (#90).
///
/// <para>
/// These are the only tests in the suite that decide for themselves whether Outlook is available,
/// rather than asking <see cref="OutlookInteropRunner"/> and skipping when it says no - because
/// <b>that answer is the thing under test</b>. Every other integration test skips itself when
/// <c>TryGetRunningApplication</c> returns false, so a machine where Outlook runs but never
/// registers in the ROT produces a fully green run that verified nothing at all. Here the presence
/// of an <c>OUTLOOK.EXE</c> process in this session is taken as ground truth and the runner is held
/// against it.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "OutlookAvailability")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookAvailabilityTests(ITestOutputHelper output)
{
    /// <summary>
    /// If Outlook is running, the runner must find it. On this machine it did not: Outlook was
    /// running with its mailbox loaded and <c>GetActiveObject</c> still returned MK_E_UNAVAILABLE,
    /// so every call reported "Outlook does not appear to be running".
    /// </summary>
    [SkippableFact]
    public void TryGetRunningApplication_FindsOutlook_WheneverAnOutlookProcessIsRunning()
    {
        Skip.IfNot(IsOutlookProcessRunning(), "Outlook is not running, so there is nothing to find.");

        bool found = OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application);

        try
        {
            Assert.True(
                found,
                "An OUTLOOK.EXE process is running in this session, but the runner reported no reachable Outlook. "
                + "That is the #90 failure: a confidently wrong answer about observable external state.");
        }
        finally
        {
            if (application != null && Marshal.IsComObject(application))
            {
                _ = Marshal.ReleaseComObject(application);
            }
        }
    }

    /// <summary>
    /// "Found it" is not enough - the reference has to be usable, and it has to be the user's
    /// mailbox rather than some half-initialised object.
    /// </summary>
    [SkippableFact]
    public void TryGetRunningApplication_ReturnsAUsableMapiSession()
    {
        Skip.IfNot(IsOutlookProcessRunning(), "Outlook is not running, so there is nothing to find.");

        Assert.True(OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application));

        OutlookInterop.NameSpace? session = null;

        try
        {
            session = application!.GetNamespace("MAPI");
            int rootFolders = session.Folders.Count;
            output.WriteLine($"MAPI session reachable; {rootFolders} root folder(s).");
            Assert.True(rootFolders > 0);
        }
        finally
        {
            if (session != null && Marshal.IsComObject(session))
            {
                _ = Marshal.FinalReleaseComObject(session);
            }

            if (application != null && Marshal.IsComObject(application))
            {
                _ = Marshal.ReleaseComObject(application);
            }
        }
    }

    /// <summary>
    /// End to end through a real command, since that is what a caller actually sees. A runner that
    /// resolves an Application but produces commands that still fail would satisfy the tests above.
    /// </summary>
    [SkippableFact]
    public void MailList_Succeeds_WheneverAnOutlookProcessIsRunning()
    {
        Skip.IfNot(IsOutlookProcessRunning(), "Outlook is not running, so there is nothing to list.");

        var result = new MailCommands().List(folder: "inbox", maxCount: 1);

        Assert.True(result.Success, result.ErrorMessage);
    }

    /// <summary>
    /// The guard that makes the fallback safe: attaching to a running Outlook is fine, starting one
    /// is not. Asserted by counting processes across a call made while Outlook is closed - the call
    /// must fail rather than conjure an Outlook window on the user's desktop.
    /// </summary>
    [SkippableFact]
    public void ResolvingOutlook_NeverStartsOutlook()
    {
        Skip.If(IsOutlookProcessRunning(), "Outlook is running; this test only means something when it is closed.");

        bool found = OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application);

        try
        {
            Assert.False(found);
            Assert.False(IsOutlookProcessRunning(), "Resolving Outlook started an Outlook process. It must never do that.");
        }
        finally
        {
            if (application != null && Marshal.IsComObject(application))
            {
                _ = Marshal.ReleaseComObject(application);
            }
        }
    }

    private static bool IsOutlookProcessRunning()
    {
        using Process current = Process.GetCurrentProcess();
        int sessionId = current.SessionId;

        foreach (Process process in Process.GetProcessesByName("OUTLOOK"))
        {
            try
            {
                if (process.SessionId == sessionId)
                {
                    return true;
                }
            }
            catch (InvalidOperationException)
            {
                // Exited between enumeration and inspection.
            }
            finally
            {
                process.Dispose();
            }
        }

        return false;
    }
}
