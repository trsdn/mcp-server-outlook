#pragma warning disable IDE0005
using System.Reflection;
using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Commands.Attachment;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.Contact;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.Tasks;
using OutlookMcp.McpServer.Tools;
using Xunit;
#pragma warning restore IDE0005

namespace OutlookMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Targeting defaults, and the generator bug that made them a lie (#9).
///
/// <para>
/// An action-dispatch tool has <b>one</b> MCP parameter per exposed name, shared by every action.
/// The generator took the first declaring method's default and applied it to all of them, so
/// <c>contact.delete</c>'s carefully chosen <c>useActiveContact = false</c> was overwritten by
/// <c>contact.read</c>'s <c>true</c> before it ever reached an LLM. Every per-action targeting
/// default in this repository was decorative on the MCP surface.
/// </para>
///
/// <para>
/// This matters because it is a confused deputy: the model chooses the verb, and the human's
/// current Outlook selection silently chooses the object. "Delete that message" then means whatever
/// happens to be highlighted when the call lands.
/// </para>
///
/// <para>
/// These are pure reflection tests over interface metadata and generated method signatures. No COM
/// object is touched, real or substituted, so they satisfy the ADR-001 exception.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ActiveTargeting")]
[Trait("RequiresOutlook", "false")]
public class ActiveTargetingDefaultTests
{
    public static TheoryData<Type, Type, string> ToolSurfaces() => new()
    {
        { typeof(IMailCommands), typeof(OutlookMailTool), "OutlookMail" },
        { typeof(IAttachmentCommands), typeof(OutlookAttachmentTool), "OutlookAttachment" },
        { typeof(ICalendarCommands), typeof(OutlookCalendarTool), "OutlookCalendar" },
        { typeof(IContactCommands), typeof(OutlookContactTool), "OutlookContact" },
        { typeof(ITaskCommands), typeof(OutlookTaskTool), "OutlookTask" },
    };

    /// <summary>
    /// The general invariant: where the Core interface's methods disagree about a parameter's
    /// default, the single merged MCP parameter must not pick a winner. It has to arrive as null so
    /// the per-action default in Core is what applies.
    /// </summary>
    [Theory]
    [MemberData(nameof(ToolSurfaces))]
    public void MergedParameters_WithDisagreeingCoreDefaults_ArriveAsNull(
        Type interfaceType,
        Type toolType,
        string toolMethodName)
    {
        var disagreeing = FindParametersWithDisagreeingDefaults(interfaceType);
        Assert.NotEmpty(disagreeing);

        var toolMethod = toolType.GetMethod(toolMethodName, BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(toolMethod);

        foreach (var (name, defaults) in disagreeing)
        {
            string snake = ToSnakeCase(name);
            var toolParameter = toolMethod!.GetParameters().FirstOrDefault(p => p.Name == snake);
            Assert.True(
                toolParameter != null,
                $"{toolType.Name} exposes no '{snake}' parameter for {interfaceType.Name}.{name}.");

            string rendered = string.Join(", ", defaults);
            Assert.True(
                Nullable.GetUnderlyingType(toolParameter!.ParameterType) != null
                    || !toolParameter.ParameterType.IsValueType,
                $"{toolType.Name}.{snake} is '{toolParameter.ParameterType.Name}', not nullable. "
                + $"{interfaceType.Name} declares conflicting defaults for it ({rendered}), so a "
                + "non-nullable MCP parameter silently forces one action's default onto all of them.");

            // The MCP schema an LLM sees comes from [DefaultValue], not from a C# default value
            // expression - the generated tool method declares none. Assert on what the SDK reads.
            var defaultAttribute = toolParameter.GetCustomAttribute<System.ComponentModel.DefaultValueAttribute>();
            Assert.True(
                defaultAttribute != null,
                $"{toolType.Name}.{snake} carries no [DefaultValue], so the MCP schema states no default at all.");

            Assert.True(
                defaultAttribute!.Value == null,
                $"{toolType.Name}.{snake} advertises a default of '{defaultAttribute.Value}' in its MCP "
                + $"schema. It must advertise null so each action's own Core default applies; "
                + $"{interfaceType.Name} declares conflicting defaults ({rendered}).");
        }
    }

    /// <summary>
    /// The specific decision this fix exists to make stick: a mutating action must not fall back to
    /// the user's current Outlook selection. Reading may; changing, deleting and sending may not.
    /// </summary>
    [Theory]
    // Mail: reads keep the fallback, mutations lose it.
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Read), "useActiveMail", true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.GetConversation), "useActiveMail", true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Export), "useActiveMail", true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Send), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Move), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Delete), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.SetReadState), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.SetFlag), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.SetCategories), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.SetSubject), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.SetBody), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.SetRecipients), "useActiveMail", false)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.RespondToMeeting), "useActiveMail", false)]
    // Drafts are inert until send, which is itself explicitly targeted and confirm-gated, so
    // replying to "the message I am looking at" stays available.
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Reply), "useActiveMail", true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.ReplyAll), "useActiveMail", true)]
    [InlineData(typeof(IMailCommands), nameof(IMailCommands.Forward), "useActiveMail", true)]
    // Attachments.
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.List), "useActiveMail", true)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.Save), "useActiveMail", true)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.Add), "useActiveMail", false)]
    [InlineData(typeof(IAttachmentCommands), nameof(IAttachmentCommands.Remove), "useActiveMail", false)]
    // Calendar, contact and task already declared the right defaults; the generator ignored them.
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.Read), "useActiveAppointment", true)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.Export), "useActiveAppointment", true)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.UpdateAppointment), "useActiveAppointment", false)]
    [InlineData(typeof(ICalendarCommands), nameof(ICalendarCommands.DeleteAppointment), "useActiveAppointment", false)]
    [InlineData(typeof(IContactCommands), nameof(IContactCommands.Read), "useActiveContact", true)]
    [InlineData(typeof(IContactCommands), nameof(IContactCommands.Update), "useActiveContact", false)]
    [InlineData(typeof(IContactCommands), nameof(IContactCommands.Delete), "useActiveContact", false)]
    [InlineData(typeof(ITaskCommands), nameof(ITaskCommands.Read), "useActiveTask", true)]
    [InlineData(typeof(ITaskCommands), nameof(ITaskCommands.Update), "useActiveTask", false)]
    [InlineData(typeof(ITaskCommands), nameof(ITaskCommands.Delete), "useActiveTask", false)]
    public void CoreAction_TargetingDefault_IsExplicitForMutations(
        Type interfaceType,
        string methodName,
        string parameterName,
        bool expected)
    {
        var method = interfaceType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance);
        Assert.NotNull(method);

        var parameter = method!.GetParameters().FirstOrDefault(p => p.Name == parameterName);
        Assert.True(parameter != null, $"{interfaceType.Name}.{methodName} has no '{parameterName}'.");
        Assert.True(parameter!.HasDefaultValue, $"{interfaceType.Name}.{methodName}.{parameterName} has no default.");

        Assert.True(
            Equals(parameter.DefaultValue, expected),
            $"{interfaceType.Name}.{methodName}: {parameterName} defaults to {parameter.DefaultValue}, "
            + $"expected {expected}. A mutating action falling back to the user's current Outlook "
            + "selection is a confused deputy: the model picks the verb, the human unknowingly picks "
            + "the object.");
    }

    /// <summary>
    /// Mirrors the generator's own camelCase-to-snake_case conversion. Duplicated rather than
    /// referenced because the generator assembly targets netstandard2.0 and is consumed as an
    /// analyzer, not as a library.
    /// </summary>
    private static string ToSnakeCase(string camelCase)
    {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < camelCase.Length; i++)
        {
            char c = camelCase[i];
            if (char.IsUpper(c))
            {
                if (i > 0)
                {
                    sb.Append('_');
                }

                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    private static List<(string Name, List<object?> Defaults)> FindParametersWithDisagreeingDefaults(Type interfaceType)
    {
        var byName = new Dictionary<string, List<object?>>(StringComparer.Ordinal);

        foreach (var method in interfaceType.GetMethods(BindingFlags.Public | BindingFlags.Instance))
        {
            if (method.GetCustomAttribute<ServiceActionAttribute>() == null)
            {
                continue;
            }

            foreach (var parameter in method.GetParameters())
            {
                if (!parameter.HasDefaultValue || parameter.Name == null)
                {
                    continue;
                }

                if (!byName.TryGetValue(parameter.Name, out var defaults))
                {
                    defaults = [];
                    byName[parameter.Name] = defaults;
                }

                if (!defaults.Contains(parameter.DefaultValue))
                {
                    defaults.Add(parameter.DefaultValue);
                }
            }
        }

        return byName
            .Where(kv => kv.Value.Count > 1)
            .Select(kv => (kv.Key, kv.Value))
            .ToList();
    }
}
