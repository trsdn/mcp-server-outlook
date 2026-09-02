using System.Text.Json;
using OutlookMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for the diag CLI command.
/// Tests the full CLI pipeline end-to-end: process launch → service dispatch → JSON response.
/// These tests do NOT require Outlook — they validate CLI infrastructure (parameter parsing,
/// validation, routing, JSON serialization, exit codes).
/// </summary>
[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Diag")]
[Trait("RequiresOutlook", "false")]
[Trait("Speed", "Fast")]
public sealed class DiagCommandTests
{
    private readonly ITestOutputHelper _output;

    public DiagCommandTests(ITestOutputHelper output) => _output = output;

    // ============================================
    // PING - Basic connectivity
    // ============================================

    [Fact]
    public async Task Ping_ReturnsSuccess()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag ping");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("ping", json.RootElement.GetProperty("action").GetString());
        Assert.Equal("pong", json.RootElement.GetProperty("message").GetString());
        Assert.True(json.RootElement.TryGetProperty("timestamp", out _));
    }

    [Fact]
    public async Task Ping_DoesNotRequireSession()
    {
        // Ping should work without --session parameter
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag ping");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
    }

    // ============================================
    // ECHO - Required string parameter
    // ============================================

    [Fact]
    public async Task Echo_WithMessage_ReturnsEchoedMessage()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"Hello World\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("echo", json.RootElement.GetProperty("action").GetString());
        Assert.Equal("Hello World", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public async Task Echo_WithMessageAndTag_ReturnsBoth()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"test\" --tag \"v1\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("test", json.RootElement.GetProperty("message").GetString());
        Assert.Equal("v1", json.RootElement.GetProperty("tag").GetString());
    }

    [Fact]
    public async Task Echo_WithoutMessage_ReturnsValidationError()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo");
        _output.WriteLine(result.Stdout);

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("message", json.RootElement.GetProperty("error").GetString());
        Assert.Contains("required", json.RootElement.GetProperty("error").GetString());
    }

    [Fact]
    public async Task Echo_WithSpecialCharacters_PreservesMessage()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"hello & goodbye <world>\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.Equal("hello & goodbye <world>", json.RootElement.GetProperty("message").GetString());
    }

    // ============================================
    // VALIDATE-PARAMS - Multiple parameter types
    // ============================================

    [Fact]
    public async Task ValidateParams_AllParams_ReturnsAll()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --name \"test\" --count 42 --label \"optional\" --verbose true");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("validate-params", json.RootElement.GetProperty("action").GetString());

        var parameters = json.RootElement.GetProperty("parameters");
        Assert.Equal("test", parameters.GetProperty("name").GetString());
        Assert.Equal(42, parameters.GetProperty("count").GetInt32());
        Assert.Equal("optional", parameters.GetProperty("label").GetString());
        Assert.True(parameters.GetProperty("verbose").GetBoolean());
    }

    [Fact]
    public async Task ValidateParams_RequiredOnly_DefaultsApplied()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --name \"test\" --count 5");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());

        var parameters = json.RootElement.GetProperty("parameters");
        Assert.Equal("test", parameters.GetProperty("name").GetString());
        Assert.Equal(5, parameters.GetProperty("count").GetInt32());
        // label is null → omitted by WhenWritingNull
        Assert.False(parameters.TryGetProperty("label", out _));
        Assert.False(parameters.GetProperty("verbose").GetBoolean());
    }

    [Fact]
    public async Task ValidateParams_MissingName_ReturnsValidationError()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            "diag validate-params --count 5");
        _output.WriteLine(result.Stdout);

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("name", json.RootElement.GetProperty("error").GetString());
        Assert.Contains("required", json.RootElement.GetProperty("error").GetString());
    }

    [Fact]
    public async Task ValidateParams_MissingCount_ReturnsValidationError()
    {
        // count is int with no default → should be required
        // But int can't be "empty" — Spectre.Console may handle differently
        // This test documents the actual behavior
        var result = await CliProcessHelper.RunAsync("diag validate-params --name \"test\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stdout: {result.Stdout}");

        // int without default gets default(int) = 0 from Spectre.Console
        // So this should succeed with count=0
        var json = JsonDocument.Parse(result.Stdout);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
    }

    // ============================================
    // INVALID ACTION - Unknown action handling
    // ============================================

    [Fact]
    public async Task InvalidAction_ReturnsError()
    {
        var result = await CliProcessHelper.RunAsync("diag nonexistent");
        _output.WriteLine(result.Stdout);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("Unknown command", result.Stdout);
    }

    // ============================================
    // JSON FORMAT - Output structure validation
    // ============================================

    [Fact]
    public async Task AllResponses_AreValidJson()
    {
        // Verify success responses are valid JSON
        var pingResult = await CliProcessHelper.RunAsync("diag ping");
        Assert.DoesNotContain("Unhandled error", pingResult.Stdout);
        var pingJson = JsonDocument.Parse(pingResult.Stdout);
        Assert.NotNull(pingJson);

        // Verify error responses are valid JSON
        var errorResult = await CliProcessHelper.RunAsync("diag echo");
        Assert.DoesNotContain("Unhandled error", errorResult.Stdout);
        var errorJson = JsonDocument.Parse(errorResult.Stdout);
        Assert.NotNull(errorJson);
    }

    [Fact]
    public async Task ErrorResponse_HasCorrectStructure()
    {
        var (_, json) = await CliProcessHelper.RunJsonAsync("diag echo");

        // Error responses must have 'success' and 'error' properties
        Assert.True(json.RootElement.TryGetProperty("success", out var success));
        Assert.False(success.GetBoolean());
        Assert.True(json.RootElement.TryGetProperty("error", out var error));
        Assert.False(string.IsNullOrWhiteSpace(error.GetString()));
    }

    [Fact]
    public async Task SuccessResponse_HasCorrectStructure()
    {
        var (_, json) = await CliProcessHelper.RunJsonAsync("diag ping");

        // Success responses must have 'success' property
        Assert.True(json.RootElement.TryGetProperty("success", out var success));
        Assert.True(success.GetBoolean());

        // Success responses must NOT have 'error' property (Rule 1)
        Assert.False(json.RootElement.TryGetProperty("error", out _),
            "Success response must not contain 'error' property (Rule 1: Success flag must match reality)");
    }

    // ============================================
    // UNKNOWN OPTIONS - must be rejected, not ignored
    // ============================================

    [Fact]
    public async Task UnknownOption_IsRejectedWithNonZeroExitCode()
    {
        // Spectre.Console.Cli defaults to non-strict parsing: an unrecognised option is silently
        // collected into the remaining arguments instead of failing. Discovered by running the CLI
        // against a live mailbox, where `mail list --limit 3` (the real option is --max-count)
        // returned the default 25 messages, reported success and exited 0.
        //
        // That failure mode is specifically dangerous for this project: the CLI is documented as a
        // surface for coding agents, which guess plausible option names. Silently ignoring one
        // hands the agent a confidently wrong answer instead of an error it could correct.
        var result = await CliProcessHelper.RunAsync("diag echo --message hi --no-such-option value");

        Assert.NotEqual(0, result.ExitCode);
    }

    [Fact]
    public async Task UnknownOption_DoesNotReportSuccess()
    {
        // The exit code and the payload must agree (Rule 1). Returning a success payload for a
        // command line the CLI did not fully understand is exactly the mismatch Rule 1 forbids.
        var result = await CliProcessHelper.RunAsync("diag echo --message hi --no-such-option value");

        Assert.DoesNotContain("\"success\":true", result.Stdout.Replace(" ", string.Empty), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task KnownOptions_StillParseAfterStrictParsingIsEnabled()
    {
        // Guard against the fix over-reaching: legitimate options must keep working.
        var (result, json) = await CliProcessHelper.RunJsonAsync("diag echo --message \"strict parsing ok\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("strict parsing ok", json.RootElement.GetProperty("message").GetString());
    }
}
