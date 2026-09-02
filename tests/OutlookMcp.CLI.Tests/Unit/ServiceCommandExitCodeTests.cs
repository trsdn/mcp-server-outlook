using OutlookMcp.CLI.Infrastructure;
using Xunit;

namespace OutlookMcp.CLI.Tests.Unit;

/// <summary>
/// REGRESSION TESTS for #63: the CLI returned exit code 0 when an operation failed.
///
/// <see cref="ServiceCommandBase{TSettings}"/> checked <c>response.Success</c>, which is a
/// <em>transport</em>-level flag meaning "the daemon replied". The operation's own outcome lives
/// in the <c>success</c> property of the JSON payload carried in <c>response.Result</c>, and it was
/// never inspected. A failed Outlook operation therefore produced <c>{"success": false, ...}</c> on
/// stdout alongside exit code 0, so every script and CI step that branched on <c>$LASTEXITCODE</c>
/// treated the failure as a success.
///
/// These are pure string/JSON tests with no COM dependency, which is the documented exception to
/// Rule 30.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "CliExitCode")]
[Trait("Speed", "Fast")]
public sealed class ServiceCommandExitCodeTests
{
    /// <summary>
    /// The core regression: a payload reporting <c>success: false</c> must map to a non-zero
    /// exit code even though the daemon replied successfully.
    /// </summary>
    [Fact]
    public void ResolveExitCode_OperationReportedFailure_ReturnsNonZero()
    {
        const string payload = """{"success":false,"errorMessage":"Folder 'Nope' not found"}""";

        Assert.NotEqual(0, ServiceCommandBase.ResolveExitCode(payload));
    }

    [Fact]
    public void ResolveExitCode_OperationReportedSuccess_ReturnsZero()
    {
        const string payload = """{"success":true,"itemCount":3}""";

        Assert.Equal(0, ServiceCommandBase.ResolveExitCode(payload));
    }

    /// <summary>
    /// The service serializes with camelCase, but a payload that reaches us as PascalCase must not
    /// be silently treated as a success.
    /// </summary>
    [Fact]
    public void ResolveExitCode_PascalCaseSuccessProperty_IsHonoured()
    {
        const string payload = """{"Success":false,"ErrorMessage":"boom"}""";

        Assert.NotEqual(0, ServiceCommandBase.ResolveExitCode(payload));
    }

    /// <summary>
    /// Many read operations return a bare array or an object with no <c>success</c> property at all.
    /// Absence of the flag is not failure - the daemon already told us the call succeeded.
    /// </summary>
    [Theory]
    [InlineData("""{"folders":[{"name":"Inbox"}]}""")]
    [InlineData("""[{"subject":"hello"}]""")]
    [InlineData("")]
    [InlineData(null)]
    public void ResolveExitCode_NoSuccessProperty_ReturnsZero(string? payload)
    {
        Assert.Equal(0, ServiceCommandBase.ResolveExitCode(payload));
    }

    /// <summary>
    /// Output that is not JSON at all must not be reported as a failure; the daemon already
    /// confirmed the operation ran, and guessing here would turn plain-text results into errors.
    /// </summary>
    [Fact]
    public void ResolveExitCode_NonJsonPayload_ReturnsZero()
    {
        Assert.Equal(0, ServiceCommandBase.ResolveExitCode("not json at all"));
    }

    /// <summary>
    /// A non-boolean <c>success</c> value is malformed rather than a failure report, and must not
    /// crash the exit-code decision.
    /// </summary>
    [Fact]
    public void ResolveExitCode_NonBooleanSuccessValue_ReturnsZero()
    {
        Assert.Equal(0, ServiceCommandBase.ResolveExitCode("""{"success":"yes"}"""));
    }
}
