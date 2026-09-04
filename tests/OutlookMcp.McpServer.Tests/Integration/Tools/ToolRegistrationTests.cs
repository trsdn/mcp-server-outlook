using System.Reflection;
using ModelContextProtocol.Server;
using Xunit;

namespace OutlookMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// The MCP server registers tools from an explicit allow-list rather than by scanning the assembly
/// (#23), so that a newly generated tool cannot silently appear in tools/list.
///
/// The cost of that choice is the opposite failure, and it shipped: `contact` was generated,
/// implemented, routed through the CLI, counted in README.md and FEATURES.md, and documented as
/// being "identical through the MCP server and the CLI" - while never being added to the allow-list.
/// It was absent from tools/list for its entire life. Nothing caught it, because every test that
/// touched the contact surface went through the Core commands or the CLI.
///
/// This test closes that gap: whatever the generator emits must be registered, or the build fails.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ToolRegistration")]
public class ToolRegistrationTests
{
    /// <summary>
    /// Every generated [McpServerToolType] class in the server assembly must be in the allow-list.
    /// </summary>
    [Fact]
    public void RegisteredToolsMatchGeneratedTools()
    {
        var generated = typeof(Program).Assembly
            .GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .Select(t => t.FullName!)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();

        Assert.NotEmpty(generated);

        var registered = Program.RegisteredToolTypes
            .Select(t => t.FullName!)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();

        var missing = generated.Except(registered, StringComparer.Ordinal).ToList();

        Assert.True(
            missing.Count == 0,
            $"These generated MCP tools are not registered in Program.RegisteredToolTypes, so they " +
            $"are absent from tools/list even though the CLI exposes them: {string.Join(", ", missing)}. " +
            "Add them to the allow-list, or delete the interface if the tool is not meant to ship.");
    }

    /// <summary>
    /// The allow-list must not name a type that is no longer generated, and must not name the same
    /// tool twice - either would mean the list has drifted from the generator.
    /// </summary>
    [Fact]
    public void RegisteredToolsAreAllRealAndDistinct()
    {
        var generated = typeof(Program).Assembly
            .GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .Select(t => t.FullName!)
            .ToHashSet(StringComparer.Ordinal);

        var registered = Program.RegisteredToolTypes.Select(t => t.FullName!).ToList();

        var stale = registered.Where(n => !generated.Contains(n)).ToList();
        Assert.True(stale.Count == 0,
            $"Registered types that no longer carry [McpServerToolType]: {string.Join(", ", stale)}");

        var duplicates = registered
            .GroupBy(n => n, StringComparer.Ordinal)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();
        Assert.True(duplicates.Count == 0,
            $"Duplicate entries in the allow-list: {string.Join(", ", duplicates)}");
    }

    /// <summary>
    /// The `task` tool exposes update and delete, so its generated tool-level destructive hint must
    /// be true. Same check as <c>DestructiveAnnotationTests</c> applies to calendar and attachment.
    /// </summary>
    [Fact]
    public void TaskTool_DeclaresDestructiveTrue()
    {
        var method = typeof(OutlookMcp.McpServer.Tools.OutlookTaskTool)
            .GetMethod("OutlookTask", BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(method);

        var attr = method!.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(attr);
        Assert.True(attr!.Destructive,
            $"task tool declares Destructive={attr.Destructive}, but it exposes update and delete.");
    }
}
