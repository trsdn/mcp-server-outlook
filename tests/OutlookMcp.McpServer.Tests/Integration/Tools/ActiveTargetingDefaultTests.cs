#pragma warning disable IDE0005
using System.Reflection;
using ModelContextProtocol.Server;
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
/// <b>The surfaces are discovered, not listed.</b> The defect being guarded here was one where the
/// declaration and the behaviour disagreed and nothing in between reported it; a guard that depends
/// on somebody remembering to extend a hand-written list has the same shape - it does not fail when
/// it falls behind, it just quietly covers less. So the surfaces come from reflection over
/// <c>[ServiceCategory]</c>, the same way <c>CoreCommandsCoverageTests</c> does it, and
/// <see cref="EveryServiceCategoryInterface_IsPairedWithAGeneratedTool"/> fails loudly if a category
/// ever appears without a matching generated tool.
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
    /// <summary>
    /// Every <c>[ServiceCategory]</c> interface in Core, paired with the MCP tool generated from it.
    /// </summary>
    public static TheoryData<string> ToolSurfaces()
    {
        var data = new TheoryData<string>();
        foreach (string category in DiscoverServiceCategories().Keys.OrderBy(k => k, StringComparer.Ordinal))
        {
            data.Add(category);
        }

        return data;
    }

    /// <summary>
    /// Guards the discovery itself. Reflection that finds nothing would make every theory below
    /// vacuously pass, which is precisely the failure this whole file exists to catch.
    /// </summary>
    [Fact]
    public void Core_ExposesAtLeastOneServiceCategory()
    {
        Assert.NotEmpty(DiscoverServiceCategories());
    }

    /// <summary>
    /// A category with no generated tool would silently drop out of the theory data below rather
    /// than failing, so the pairing is asserted separately and by name.
    /// </summary>
    [Theory]
    [MemberData(nameof(ToolSurfaces))]
    public void EveryServiceCategoryInterface_IsPairedWithAGeneratedTool(string category)
    {
        Type interfaceType = DiscoverServiceCategories()[category];

        Assert.True(
            TryResolveGeneratedTool(category, out Type? toolType, out MethodInfo? toolMethod),
            $"Category '{category}' ({interfaceType.Name}) has no generated MCP tool. Expected a type "
            + $"'OutlookMcp.McpServer.Tools.Outlook{ToPascal(category)}Tool' with a public static "
            + $"'Outlook{ToPascal(category)}' method carrying [McpServerTool]. If the generator's "
            + "naming convention changed, update this test rather than deleting it - it is what stops "
            + "a new tool surface acquiring no targeting-default guard at all.");

        Assert.NotNull(toolType);
        Assert.NotNull(toolMethod);
    }

    /// <summary>
    /// The anti-vacuity guard, applied across the whole surface rather than per category.
    ///
    /// <para>
    /// It cannot be per category: <c>application</c> and <c>folder</c> legitimately declare no
    /// parameter whose defaults disagree, so requiring one of every tool would fail on a correct
    /// repository. What must never happen is <i>no</i> tool having one, because that would mean the
    /// theory below is asserting nothing anywhere - either the invariant genuinely no longer applies
    /// and this file should go, or reflection has quietly stopped seeing the defaults.
    /// </para>
    /// </summary>
    [Fact]
    public void AtLeastOneSurface_HasAParameterWhoseCoreDefaultsDisagree()
    {
        var withDisagreements = DiscoverServiceCategories()
            .Where(kv => FindParametersWithDisagreeingDefaults(kv.Value).Count > 0)
            .Select(kv => kv.Key)
            .ToList();

        Assert.NotEmpty(withDisagreements);
    }

    /// <summary>
    /// The general invariant: where the Core interface's methods disagree about a parameter's
    /// default, the single merged MCP parameter must not pick a winner. It has to arrive as null so
    /// the per-action default in Core is what applies.
    /// </summary>
    [Theory]
    [MemberData(nameof(ToolSurfaces))]
    public void MergedParameters_WithDisagreeingCoreDefaults_ArriveAsNull(string category)
    {
        Type interfaceType = DiscoverServiceCategories()[category];

        Assert.True(
            TryResolveGeneratedTool(category, out Type? toolType, out MethodInfo? toolMethod),
            $"Category '{category}' has no generated MCP tool to check.");

        var disagreeing = FindParametersWithDisagreeingDefaults(interfaceType);

        foreach (var (name, defaults) in disagreeing)
        {
            string snake = ToSnakeCase(name);
            var toolParameter = toolMethod!.GetParameters().FirstOrDefault(p => p.Name == snake);
            Assert.True(
                toolParameter != null,
                $"{toolType!.Name} exposes no '{snake}' parameter for {interfaceType.Name}.{name}.");

            string rendered = string.Join(", ", defaults);
            Assert.True(
                Nullable.GetUnderlyingType(toolParameter!.ParameterType) != null
                    || !toolParameter.ParameterType.IsValueType,
                $"{toolType!.Name}.{snake} is '{toolParameter.ParameterType.Name}', not nullable. "
                + $"{interfaceType.Name} declares conflicting defaults for it ({rendered}), so a "
                + "non-nullable MCP parameter silently forces one action's default onto all of them.");

            // The MCP schema an LLM sees comes from [DefaultValue], not from a C# default value
            // expression - the generated tool method declares none. Assert on what the SDK reads.
            var defaultAttribute = toolParameter.GetCustomAttribute<System.ComponentModel.DefaultValueAttribute>();
            Assert.True(
                defaultAttribute != null,
                $"{toolType!.Name}.{snake} carries no [DefaultValue], so the MCP schema states no default at all.");

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
    /// Every <c>[ServiceCategory]</c> interface in Core, keyed by its category name. Reflection
    /// rather than a list, for the reason given on the class.
    /// </summary>
    private static Dictionary<string, Type> DiscoverServiceCategories()
    {
        return typeof(IMailCommands).Assembly
            .GetTypes()
            .Where(t => t.IsInterface)
            .Select(t => (Type: t, Attr: t.GetCustomAttribute<ServiceCategoryAttribute>()))
            .Where(x => x.Attr is not null)
            .ToDictionary(x => x.Attr!.Category, x => x.Type, StringComparer.Ordinal);
    }

    /// <summary>
    /// Resolves the tool the MCP generator produces for a category: type
    /// <c>Outlook{Pascal}Tool</c> with a public static <c>Outlook{Pascal}</c> method carrying
    /// <c>[McpServerTool]</c>.
    /// </summary>
    /// <remarks>
    /// The <c>[McpServerTool]</c> check is not decoration. A method of the right name that the SDK
    /// does not treat as a tool is not a surface an LLM can reach, and pairing against it would
    /// report coverage this file does not have.
    /// </remarks>
    private static bool TryResolveGeneratedTool(string category, out Type? toolType, out MethodInfo? toolMethod)
    {
        string pascal = ToPascal(category);

        toolType = typeof(OutlookToolsBase).Assembly
            .GetType($"OutlookMcp.McpServer.Tools.Outlook{pascal}Tool");

        toolMethod = toolType?
            .GetMethod($"Outlook{pascal}", BindingFlags.Public | BindingFlags.Static);

        if (toolMethod?.GetCustomAttribute<McpServerToolAttribute>() == null)
        {
            toolMethod = null;
        }

        return toolType != null && toolMethod != null;
    }

    /// <summary>
    /// The Pascal form the generators use for a category, honouring an explicit
    /// <c>[ServiceCategory(PascalName = ...)]</c> where one is given - the same resolution
    /// <c>CoreCommandsCoverageTests</c> performs.
    /// </summary>
    private static string ToPascal(string category)
    {
        var attr = DiscoverServiceCategories()[category].GetCustomAttribute<ServiceCategoryAttribute>();
        if (!string.IsNullOrEmpty(attr?.PascalName))
        {
            return attr!.PascalName!;
        }

        return string.Concat(char.ToUpperInvariant(category[0]), category[1..]);
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
