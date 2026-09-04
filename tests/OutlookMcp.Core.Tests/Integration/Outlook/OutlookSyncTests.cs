using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Commands.Sync;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Send/Receive (sync) operations against the running Outlook, backed by
/// <c>Namespace.SyncObjects</c> (#15).
///
/// <para>
/// These tests are written not to assume the owner's profile is in Cached Exchange mode. A pure
/// Online profile legitimately has zero Send/Receive groups, and a test that only passed with groups
/// present would silently stop testing anything on such a profile. What is asserted instead is the
/// property that must hold either way: the connection mode is reported, the group count and the group
/// list agree, and an unknown group name is rejected with the available names.
/// </para>
///
/// <para>
/// <c>send-receive</c> is asynchronous. The only real-sync assertion made here is that starting all
/// groups succeeds and reports which groups it started; the operation is the same routine background
/// synchronisation Outlook performs on its own timer, so it is safe on a shared mailbox.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Sync")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookSyncTests(ITestOutputHelper output)
{
    [SkippableFact]
    public void ListGroups_ReportsConnectionModeAndAgreesWithCount()
    {
        EnsureOutlookAvailable();

        var result = new SyncCommands().ListGroups();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.ExchangeConnectionMode),
            "The Exchange connection mode should always be reported.");
        Assert.False(string.IsNullOrWhiteSpace(result.CacheMode));
        Assert.Equal(result.Count, result.Groups.Count);

        output.WriteLine(
            $"connectionMode={result.ExchangeConnectionMode} cacheMode={result.CacheMode} "
            + $"groups={result.Count}");
        foreach (var group in result.Groups)
        {
            Assert.False(string.IsNullOrWhiteSpace(group.Name), "A sync group arrived without a name.");
            output.WriteLine($"  group: {group.Name}");
        }
    }

    [SkippableFact]
    public void SendReceive_UnknownGroup_FailsAndListsAvailableGroups()
    {
        EnsureOutlookAvailable();

        var groups = new SyncCommands().ListGroups();
        Assert.True(groups.Success, groups.ErrorMessage);
        Skip.If(groups.Count == 0,
            "This profile has no Send/Receive groups (pure Online mode), so the unknown-group path "
            + "cannot be exercised.");

        var result = new SyncCommands().SendReceive("__no_such_group_" + Guid.NewGuid().ToString("N"));

        Assert.False(result.Success);
        Assert.False(result.Started);
        Assert.NotNull(result.ErrorMessage);
        output.WriteLine(result.ErrorMessage);
    }

    [SkippableFact]
    public void SendReceive_AllGroups_StartsSyncAndReportsItAsAsynchronous()
    {
        EnsureOutlookAvailable();

        var groups = new SyncCommands().ListGroups();
        Assert.True(groups.Success, groups.ErrorMessage);
        Skip.If(groups.Count == 0,
            "This profile has no Send/Receive groups (pure Online mode), so there is nothing to start.");

        var result = new SyncCommands().SendReceive();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Started, "Starting all groups should report at least one started group.");
        Assert.NotEmpty(result.StartedGroups);
        Assert.NotNull(result.Note);
        output.WriteLine($"started: {string.Join(", ", result.StartedGroups)}");
        output.WriteLine($"note: {result.Note}");
    }

    [SkippableFact]
    public void SendReceive_ReportsOnlineModeHonestlyWhenNoGroupsExist()
    {
        EnsureOutlookAvailable();

        var groups = new SyncCommands().ListGroups();
        Assert.True(groups.Success, groups.ErrorMessage);
        Skip.If(groups.Count != 0,
            "This profile has Send/Receive groups, so the empty-collection path does not apply here.");

        var result = new SyncCommands().SendReceive();

        // No groups is not an error: the call succeeds, reports that nothing was started, and says why.
        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.Started);
        Assert.NotNull(result.Note);
        output.WriteLine(result.Note);
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
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
