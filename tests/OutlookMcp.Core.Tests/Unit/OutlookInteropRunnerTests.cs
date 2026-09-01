using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for <see cref="OutlookInteropRunner.IsObjectModelGuardDenial"/>, a pure HRESULT
/// classification helper with zero COM dependency (Rule 30's exception for algorithmic
/// utilities). Covers #30: Object Model Guard denials must be distinguished from other COM
/// failures rather than silently swallowed as ordinary errors.
/// </summary>
public class OutlookInteropRunnerTests
{
    // COMException is used here as intentional test fixture data (CA2201 suppressed), not
    // thrown/rethrown as a real runtime error.
#pragma warning disable CA2201
    [Fact]
    public void IsObjectModelGuardDenial_WithEAbortHResult_ReturnsTrue()
    {
        const int E_ABORT = unchecked((int)0x80004004);
        var ex = new COMException("Operation aborted", E_ABORT);

        Assert.True(OutlookInteropRunner.IsObjectModelGuardDenial(ex));
    }

    [Theory]
    [InlineData(unchecked((int)0x80070005))] // E_ACCESSDENIED
    [InlineData(unchecked((int)0x800401FD))] // CO_E_CLASSNOTREG-ish placeholder
    [InlineData(0)]
    public void IsObjectModelGuardDenial_WithOtherHResults_ReturnsFalse(int hresult)
    {
        var ex = new COMException("Some other COM failure", hresult);

        Assert.False(OutlookInteropRunner.IsObjectModelGuardDenial(ex));
    }
#pragma warning restore CA2201
}
