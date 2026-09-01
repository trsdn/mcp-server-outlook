// Explicit usings retained; pragma used to suppress IDE0005 for clarity in reflection-heavy test
#pragma warning disable IDE0005
using System.Reflection;
using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Commands.Attachment;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.McpServer.Tools;
using ModelContextProtocol.Server;
using Xunit;
#pragma warning restore IDE0005

namespace OutlookMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Regression tests for issue #18: MCP `destructiveHint` must match what each action actually
/// does. Action-dispatch tools (one MCP tool, many actions) mix read-only and mutating actions,
/// so the tool-level [McpServerTool(Destructive=...)] hint must be true whenever ANY exposed
/// action mutates Outlook state, and per-[ServiceAction] overrides must resolve correctly.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "DestructiveAnnotations")]
[Trait("RequiresPowerPoint", "false")]
public class DestructiveAnnotationTests
{
    [Theory]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.List), false)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.Read), false)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.CreateAppointment), true)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.UpdateAppointment), true)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.DeleteAppointment), true)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.List), false)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.Save), false)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.Add), true)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.Remove), true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.ReadActive), false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Read), false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.List), false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Search), false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Send), true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Delete), true)]
    public void ServiceAction_DestructiveClassification_MatchesActualBehavior(Type interfaceType, string methodName, bool expectedDestructive)
    {
        var method = interfaceType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance);
        Assert.NotNull(method);

        var toolAttr = interfaceType.GetCustomAttribute<McpToolAttribute>();
        Assert.NotNull(toolAttr);

        var actionAttr = method!.GetCustomAttribute<ServiceActionAttribute>();
        Assert.NotNull(actionAttr);

        // Resolve exactly as the generator does: per-action [ServiceAction(Destructive=...)]
        // override if the named argument was actually supplied, otherwise the tool-level default.
        var actionAttrData = method.GetCustomAttributesData()
            .First(a => a.AttributeType == typeof(ServiceActionAttribute));
        var hasDestructiveOverride = actionAttrData.NamedArguments!
            .Any(na => na.MemberName == nameof(ServiceActionAttribute.Destructive));

        var resolved = hasDestructiveOverride ? actionAttr!.Destructive : toolAttr!.Destructive;

        Assert.True(resolved == expectedDestructive,
            $"{interfaceType.Name}.{methodName}: resolved Destructive={resolved}, expected {expectedDestructive}. " +
            "This action's real-world behavior no longer matches its MCP destructiveHint annotation.");
    }

    /// <summary>
    /// The generated `calendar` tool must NOT under-declare Destructive=false: it exposes
    /// delete-appointment and update-appointment, both mutating.
    /// </summary>
    [Fact]
    public void CalendarTool_DeclaresDestructiveTrue()
    {
        AssertGeneratedToolDestructive(typeof(PptCalendarTool), "PptCalendar", expected: true);
    }

    /// <summary>
    /// The generated `attachment` tool must NOT under-declare Destructive=false: it exposes
    /// add and remove, both mutating.
    /// </summary>
    [Fact]
    public void AttachmentTool_DeclaresDestructiveTrue()
    {
        AssertGeneratedToolDestructive(typeof(PptAttachmentTool), "PptAttachment", expected: true);
    }

    private static void AssertGeneratedToolDestructive(Type toolType, string methodName, bool expected)
    {
        var method = toolType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(method);

        var attr = method!.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(attr);
        Assert.True(attr!.Destructive == expected,
            $"{toolType.Name}.{methodName}: generated [McpServerTool(Destructive={attr.Destructive})] does not match expected {expected}.");
    }
}
