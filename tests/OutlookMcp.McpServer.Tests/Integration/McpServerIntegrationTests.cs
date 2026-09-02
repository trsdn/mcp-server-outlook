// Copyright (c) Sbroenne.
// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OutlookMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

// Avoid namespace conflict: McpServer is both a type and namespace
using Server = ModelContextProtocol.Server;

namespace OutlookMcp.McpServer.Tests.Integration;

/// <summary>
/// Integration tests that exercise the full MCP protocol using in-memory transport.
/// These tests use the official MCP SDK client to connect to our server, ensuring:
/// - DI pipeline is correctly configured
/// - Tool discovery via the explicit Outlook-only allow-list in Program.cs works (#23)
/// - Tool schemas are correctly generated
/// - Tools execute properly through the MCP protocol
///
/// This is the CORRECT way to test MCP servers - using the SDK's client to verify
/// the actual protocol behavior, not reflection or direct method calls.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public class McpServerIntegrationTests(ITestOutputHelper output) : IAsyncLifetime, IAsyncDisposable
{
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private Server.McpServer? _server;
    private McpClient? _client;
    private IServiceProvider? _serviceProvider;
    private Task? _serverTask;

    /// <summary>
    /// Expected tool names from our assembly - the source of truth.
    /// Program.cs registers only these five Outlook tools via an explicit allow-list (#23); the
    /// assembly still contains 33 generated legacy PowerPoint tool types plus the hand-written
    /// "file" tool, but they are deliberately not registered with the MCP server and so must not
    /// appear here or in tools/list. This test is the regression guard for that allow-list.
    /// </summary>
    private static readonly HashSet<string> ExpectedToolNames =
    [
        "application",
        "attachment",
        "calendar",
        "folder",
        "mail",
    ];

    /// <summary>
    /// Setup: Create MCP server with DI and connect client via in-memory pipes.
    /// This exercises the exact same code path as Program.cs.
    /// </summary>
    public async Task InitializeAsync()
    {
        // Build the server with DI - same pattern as Program.cs
        var services = new ServiceCollection();
        services.AddLogging(builder => builder.AddDebug().SetMinimumLevel(LogLevel.Debug));

        // Add MCP server with tools using stream transport for testing
        services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new() { Name = "OutlookMcp-Test", Version = "1.0.0" };
                options.ServerInstructions = "Test server for integration tests";
            })
            .WithStreamServerTransport(
                _clientToServerPipe.Reader.AsStream(),
                _serverToClientPipe.Writer.AsStream())
            // Mirror Program.cs's explicit Outlook-only allow-list (#23) instead of
            // WithToolsFromAssembly(), so this test genuinely exercises (and guards) the same
            // registration surface real clients see.
            .WithTools(
            [
                typeof(PptApplicationTool),
                typeof(PptAttachmentTool),
                typeof(PptCalendarTool),
                typeof(PptFolderTool),
                typeof(PptMailTool),
            ])
            .WithPrompts([typeof(OutlookMcp.McpServer.Prompts.PptSkillPrompts)]);

        _serviceProvider = services.BuildServiceProvider(validateScopes: true);

        // Get the server and start it
        _server = _serviceProvider.GetRequiredService<Server.McpServer>();
        _serverTask = _server.RunAsync(_cts.Token);

        // Create client connected to the server via pipes
        _client = await McpClient.CreateAsync(
            new StreamClientTransport(
                serverInput: _clientToServerPipe.Writer.AsStream(),
                serverOutput: _serverToClientPipe.Reader.AsStream()),
            clientOptions: new McpClientOptions
            {
                ClientInfo = new() { Name = "TestClient", Version = "1.0.0" }
            },
            cancellationToken: _cts.Token);

        output.WriteLine($"✓ Connected to server: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");
    }

    public async Task DisposeAsync()
    {
        await DisposeAsyncCore();
    }

    // Explicit IAsyncDisposable implementation to satisfy CA1001 analyzer
    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        await _cts.CancelAsync();

        _clientToServerPipe.Writer.Complete();
        _serverToClientPipe.Writer.Complete();

        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        if (_serverTask != null)
        {
            try
            {
                await _serverTask;
            }
            catch (OperationCanceledException)
            {
                // Expected during shutdown
            }
        }

        if (_serviceProvider is IAsyncDisposable asyncDisposable)
        {
            await asyncDisposable.DisposeAsync();
        }
        else if (_serviceProvider is IDisposable disposable)
        {
            disposable.Dispose();
        }

        _cts.Dispose();
    }

    /// <summary>
    /// Canonical MCP smoke test used by pre-commit.
    /// Verifies that all expected tools are discoverable through the real MCP protocol surface.
    /// This is THE definitive test - it uses client.ListToolsAsync() which exercises:
    /// - DI pipeline
    /// - Explicit Outlook-only tool allow-list (#23) discovery
    /// - MCP protocol serialization
    /// - Tool schema generation
    /// </summary>
    [Fact]
    public async Task SmokeTest_AllTools_E2EWorkflow()
    {
        output.WriteLine("=== TOOL DISCOVERY VIA MCP PROTOCOL ===\n");

        // Act - Use the REAL MCP protocol to list tools
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        // Assert - Verify count
        output.WriteLine($"Discovered {tools.Count} tools via MCP protocol:\n");

        foreach (var tool in tools.OrderBy(t => t.Name))
        {
            var descPreview = tool.Description?.Length > 60 ? tool.Description[..60] + "..." : tool.Description;
            output.WriteLine($"  • {tool.Name}: {descPreview}");
        }

        Assert.Equal(ExpectedToolNames.Count, tools.Count);

        // Verify all expected tools are present
        var actualToolNames = tools.Select(t => t.Name).ToHashSet();

        var missingTools = ExpectedToolNames.Except(actualToolNames).ToList();
        if (missingTools.Count > 0)
        {
            output.WriteLine($"\n❌ Missing tools: {string.Join(", ", missingTools)}");
        }
        Assert.Empty(missingTools);

        var unexpectedTools = actualToolNames.Except(ExpectedToolNames).ToList();
        if (unexpectedTools.Count > 0)
        {
            output.WriteLine($"\n❌ Unexpected tools: {string.Join(", ", unexpectedTools)}");
        }
        Assert.Empty(unexpectedTools);

        output.WriteLine($"\n✓ All {ExpectedToolNames.Count} tools discovered successfully via MCP protocol");
    }

    /// <summary>
    /// Tests that each tool has proper schema (parameters, descriptions).
    /// </summary>
    [Fact]
    public async Task ListTools_AllToolsHaveValidSchema()
    {
        output.WriteLine("=== TOOL SCHEMA VALIDATION ===\n");

        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        foreach (var tool in tools)
        {
            // Every tool must have a name
            Assert.False(string.IsNullOrEmpty(tool.Name), "Tool has empty name");

            // Every tool should have a description
            Assert.False(string.IsNullOrEmpty(tool.Description), $"Tool {tool.Name} has no description");

            // McpClientTool implements AIFunction which has Parameters property
            // The SDK generates schema from tool methods

            output.WriteLine($"✓ {tool.Name}: Has description ({tool.Description?.Length} chars)");
        }

        output.WriteLine($"\n✓ All {tools.Count} tools have valid schemas");
    }

    /// <summary>
    /// Tests that the application tool's get-status action works via MCP protocol.
    /// This exercises the complete tool invocation path for an Outlook tool (#23: the legacy
    /// PowerPoint "file" tool this test previously called is no longer registered).
    /// </summary>
    [Fact]
    public async Task CallTool_ApplicationGetStatus_ReturnsSuccess()
    {
        output.WriteLine("=== TOOL INVOCATION VIA MCP PROTOCOL ===\n");

        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "get-status"
        };

        // Act - Call tool via MCP protocol
        var result = await _client!.CallToolAsync(
            "application",
            arguments,
            cancellationToken: _cts.Token);

        // Assert
        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        // Get text content - need to cast from ContentBlock base class
        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        var textPreview = textBlock.Text.Length > 200 ? textBlock.Text[..200] + "..." : textBlock.Text;
        output.WriteLine($"Tool response: {textPreview}");

        output.WriteLine("\n✓ application get-status action executed successfully via MCP protocol");
    }

    /// <summary>
    /// Tests that server information is correctly exposed via MCP protocol.
    /// </summary>
    [Fact]
    public async Task ServerInfo_ReturnsCorrectInformation()
    {
        output.WriteLine("=== SERVER INFO VIA MCP PROTOCOL ===\n");

        // Act - Server info is available after connection
        var serverInfo = _client!.ServerInfo;
        var serverInstructions = _client.ServerInstructions;

        // Assert
        Assert.NotNull(serverInfo);
        Assert.Equal("OutlookMcp-Test", serverInfo.Name);
        Assert.Equal("1.0.0", serverInfo.Version);
        Assert.Equal("Test server for integration tests", serverInstructions);

        output.WriteLine($"Server Name: {serverInfo.Name}");
        output.WriteLine($"Server Version: {serverInfo.Version}");
        output.WriteLine($"Server Instructions: {serverInstructions}");

        output.WriteLine("\n✓ Server info correctly exposed via MCP protocol");
        await Task.CompletedTask; // Satisfy async requirement
    }

    /// <summary>
    /// Tests that all tools can be discovered and iterated via ListToolsAsync.
    /// Note: SDK 0.5.0+ replaced EnumerateToolsAsync with ListToolsAsync.
    /// </summary>
    [Fact]
    public async Task ListTools_CanIterateAllTools()
    {
        output.WriteLine("=== TOOL ITERATION ===\n");

        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var toolCount = 0;
        foreach (var tool in tools)
        {
            toolCount++;
            output.WriteLine($"  Discovered: {tool.Name}");
        }

        Assert.Equal(ExpectedToolNames.Count, toolCount);

        output.WriteLine($"\n✓ Iterated {toolCount} tools");
    }

    /// <summary>
    /// Tests that server capabilities include tools.
    /// </summary>
    [Fact]
    public void ServerCapabilities_IncludesTools()
    {
        output.WriteLine("=== SERVER CAPABILITIES ===\n");

        var capabilities = _client!.ServerCapabilities;

        Assert.NotNull(capabilities);
        Assert.NotNull(capabilities.Tools);

        output.WriteLine($"✓ Tools capability: {capabilities.Tools != null}");
        output.WriteLine($"✓ ListChanged: {capabilities.Tools?.ListChanged}");

        output.WriteLine("\n✓ Server capabilities correctly exposed");
    }

    /// <summary>
    /// Regression guard for #23's prompts/list acceptance criterion: only Outlook-relevant prompts
    /// should be discoverable. Every `.md` under `skills/shared/` becomes an `[McpServerPrompt]`, so
    /// this also guards against a PowerPoint-era doc being reintroduced there.
    /// </summary>
    [Fact]
    public async Task ListPrompts_ReturnsOnlyOutlookPrompts()
    {
        output.WriteLine("=== PROMPT DISCOVERY VIA MCP PROTOCOL ===\n");

        var expectedPromptNames = new HashSet<string>
        {
            "behavioral_rules_guide",
            "outlook_workflows_guide",
        };

        var prompts = await _client!.ListPromptsAsync(cancellationToken: _cts.Token);
        var actualPromptNames = prompts.Select(p => p.Name).ToHashSet();

        foreach (var prompt in prompts.OrderBy(p => p.Name))
        {
            output.WriteLine($"  • {prompt.Name}");
        }

        var missingPrompts = expectedPromptNames.Except(actualPromptNames).ToList();
        Assert.Empty(missingPrompts);

        var unexpectedPrompts = actualPromptNames.Except(expectedPromptNames).ToList();
        if (unexpectedPrompts.Count > 0)
        {
            output.WriteLine($"\n❌ Unexpected prompts (likely leaked PowerPoint prompts): {string.Join(", ", unexpectedPrompts)}");
        }
        Assert.Empty(unexpectedPrompts);

        output.WriteLine($"\n✓ All {expectedPromptNames.Count} Outlook prompts discovered, no legacy PowerPoint prompts leaked");
    }
}




