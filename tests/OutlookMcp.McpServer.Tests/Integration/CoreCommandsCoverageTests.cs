using System.Reflection;
using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Generated;
using Xunit;

namespace OutlookMcp.McpServer.Tests.Integration;

/// <summary>
/// CRITICAL: Automated verification that every Core command method is exposed via a generated
/// action, and that every generated action maps to an action string.
/// </summary>
/// <remarks>
/// This suite is driven by reflection over the <c>[ServiceCategory]</c> interfaces in
/// <c>OutlookMcp.Core</c> rather than a hand-maintained list of domains. The previous
/// hand-maintained version silently omitted <c>ICalendarCommands</c> entirely, so the Calendar
/// domain had no coverage assertions at all. Deriving the domain set from the assembly means a
/// newly added <c>[ServiceCategory]</c> interface is covered automatically, and a domain cannot
/// be forgotten.
/// </remarks>
public class CoreCommandsCoverageTests
{
    /// <summary>
    /// All <c>[ServiceCategory]</c> interfaces in Core, as xUnit theory data.
    /// </summary>
    public static TheoryData<string> ServiceCategories()
    {
        var data = new TheoryData<string>();
        foreach (var category in GetServiceCategoryInterfaces().Keys.OrderBy(k => k, StringComparer.Ordinal))
        {
            data.Add(category);
        }

        return data;
    }

    [Fact]
    public void Core_ExposesAtLeastOneServiceCategory()
    {
        // Guards against the reflection above silently finding nothing, which would make every
        // [Theory] below vacuously pass.
        Assert.NotEmpty(GetServiceCategoryInterfaces());
    }

    [Theory]
    [MemberData(nameof(ServiceCategories))]
    public void EveryCategory_HasAnActionEnum(string category)
    {
        var enumType = GetActionEnumType(category);

        Assert.True(enumType is not null,
            $"Category '{category}' has no corresponding action enum in OutlookMcp.Generated. " +
            $"Expected a type named '{ToPascal(category)}Action'.");
    }

    [Theory]
    [MemberData(nameof(ServiceCategories))]
    public void EveryCategory_HasEnumValueForEveryServiceActionMethod(string category)
    {
        var interfaceType = GetServiceCategoryInterfaces()[category];
        var enumType = GetActionEnumType(category);
        Assert.True(enumType is not null, $"Category '{category}' has no action enum.");

        var coreMethodCount = GetServiceActionMethodCount(interfaceType);
        var enumValueCount = Enum.GetValues(enumType!).Length;

        Assert.True(enumValueCount >= coreMethodCount,
            $"{interfaceType.Name} has {coreMethodCount} [ServiceAction] methods but " +
            $"{enumType!.Name} has only {enumValueCount} enum values.");
    }

    [Theory]
    [MemberData(nameof(ServiceCategories))]
    public void EveryCategory_MapsAllEnumValuesToActionStrings(string category)
    {
        var enumType = GetActionEnumType(category);
        Assert.True(enumType is not null, $"Category '{category}' has no action enum.");

        var toActionString = GetToActionStringMethod(category, enumType!);
        Assert.True(toActionString is not null,
            $"ServiceRegistry.{ToPascal(category)}.ToActionString({enumType!.Name}) was not found.");

        // Rule 15: every enum value must have a mapping. A missing case throws instead of
        // returning JSON, which surfaces to MCP clients as an unhandled invocation error.
        foreach (var action in Enum.GetValues(enumType!))
        {
            string? mapped = null;
            var exception = Record.Exception(() =>
                mapped = (string?)toActionString!.Invoke(null, [action]));

            Assert.True(exception is null,
                $"ServiceRegistry.{ToPascal(category)}.ToActionString({enumType!.Name}.{action}) threw: " +
                $"{exception?.InnerException?.Message ?? exception?.Message}");
            Assert.False(string.IsNullOrEmpty(mapped),
                $"ServiceRegistry.{ToPascal(category)}.ToActionString({enumType!.Name}.{action}) returned empty.");
        }
    }

    // ── Reflection helpers ───────────────────────────────────

    private static Dictionary<string, Type> GetServiceCategoryInterfaces()
    {
        return typeof(IMailCommands).Assembly
            .GetTypes()
            .Where(t => t.IsInterface)
            .Select(t => (Type: t, Attr: t.GetCustomAttribute<ServiceCategoryAttribute>()))
            .Where(x => x.Attr is not null)
            .ToDictionary(x => x.Attr!.Category, x => x.Type, StringComparer.Ordinal);
    }

    private static Type? GetActionEnumType(string category)
    {
        return typeof(ServiceRegistry).Assembly
            .GetType($"OutlookMcp.Generated.{ToPascal(category)}Action");
    }

    private static MethodInfo? GetToActionStringMethod(string category, Type enumType)
    {
        var registry = typeof(ServiceRegistry)
            .GetNestedType(ToPascal(category), BindingFlags.Public | BindingFlags.Static);

        return registry?
            .GetMethod("ToActionString", BindingFlags.Public | BindingFlags.Static, [enumType]);
    }

    private static string ToPascal(string category)
    {
        var attr = GetServiceCategoryInterfaces()[category].GetCustomAttribute<ServiceCategoryAttribute>();
        if (!string.IsNullOrEmpty(attr?.PascalName))
        {
            return attr!.PascalName!;
        }

        return string.Concat(char.ToUpperInvariant(category[0]), category[1..]);
    }

    /// <summary>
    /// Helper: Counts methods with [ServiceAction] attribute in an interface.
    /// </summary>
    private static int GetServiceActionMethodCount(Type interfaceType)
    {
        return interfaceType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.GetCustomAttributes()
                .Any(a => a.GetType().Name == "ServiceActionAttribute"))
            .Select(m => m.Name)
            .Distinct(StringComparer.Ordinal)
            .Count();
    }
}
