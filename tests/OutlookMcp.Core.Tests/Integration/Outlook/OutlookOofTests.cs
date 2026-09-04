using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Oof;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Read-only out-of-office (automatic replies) status against the running Outlook, backed by the
/// <c>PR_OOF_STATE</c> store property read through <c>Store.PropertyAccessor</c> (#15).
///
/// <para>
/// The test does not assume the owner is or is not out of office - it never toggles a real user's
/// automatic replies. It asserts the property that must hold either way: on an Exchange mailbox the
/// on/off flag is readable as a real boolean, the store is identified, and a note is always present.
/// On a non-Exchange store the call still succeeds but reports the feature as unsupported, and the
/// test skips the boolean assertion rather than passing vacuously.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Oof")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookOofTests(ITestOutputHelper output)
{
    [SkippableFact]
    public void GetStatus_ReadsAutomaticRepliesStateFromTheRunningMailbox()
    {
        EnsureOutlookAvailable();

        var result = new OofCommands().GetStatus();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.StoreDisplayName),
            "The default store should always be identified by name.");
        Assert.False(string.IsNullOrWhiteSpace(result.ExchangeStoreType));
        Assert.NotNull(result.Note);

        output.WriteLine(
            $"store='{result.StoreDisplayName}' type={result.ExchangeStoreType} "
            + $"supported={result.IsSupported} enabled={result.IsOutOfOfficeEnabled?.ToString() ?? "null"}");
        output.WriteLine($"note: {result.Note}");

        Skip.If(!result.IsSupported,
            "The default store is not an Exchange mailbox, so out-of-office does not apply here.");

        // On an Exchange mailbox the on/off flag must be a real boolean, never null - that is the whole
        // point of the feature and guards against a vacuous pass.
        Assert.True(result.IsOutOfOfficeEnabled.HasValue,
            "On an Exchange store the out-of-office on/off state must be readable.");
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
