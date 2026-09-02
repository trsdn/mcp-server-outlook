using System.Reflection;
using System.Text.Json;
using OutlookMcp.CLI.Commands;
using OutlookMcp.Generated;
using Spectre.Console.Cli;
using Xunit;

namespace OutlookMcp.CLI.Tests.Unit;

/// <summary>
/// Verifies that the generated <c>ServiceRegistry.{Category}.ValidActions</c> arrays stay in sync
/// with the corresponding <c>ToActionString</c> mappings, and that <see cref="ListActionsCommand"/>
/// surfaces every generated CLI category.
/// </summary>
/// <remarks>
/// These are pure reflection assertions over generated metadata with no COM dependency, which is the
/// documented exception to the integration-tests-only rule. The category list is discovered from
/// <c>_CliCategoryMetadata</c> rather than hand-maintained, so adding or removing a
/// <c>[ServiceCategory]</c> interface cannot silently drop coverage.
/// </remarks>
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ActionValidation")]
[Trait("Speed", "Fast")]
public sealed class ActionValidatorTests
{
    /// <summary>Every generated ServiceRegistry category type, discovered at runtime.</summary>
    public static TheoryData<string> RegistryCategoryNames
    {
        get
        {
            var data = new TheoryData<string>();
            foreach (var (cliCommandName, _, _) in _CliCategoryMetadata.Categories)
            {
                data.Add(cliCommandName);
            }

            return data;
        }
    }

    [Fact]
    public void Generator_EmitsAtLeastOneCliCategory()
    {
        // Guards against the theories below passing vacuously if generation breaks.
        Assert.NotEmpty(_CliCategoryMetadata.Categories);
    }

    [Theory]
    [MemberData(nameof(RegistryCategoryNames))]
    public void ValidActions_MatchesToActionStringMapping(string cliCommandName)
    {
        var registryType = ResolveRegistryType(cliCommandName);

        var expected = GetExpectedActions(registryType);
        var actual = GetActualActions(registryType);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ListActionsCommand_AllCommands_ReturnsEveryGeneratedCategory()
    {
        var command = new ListActionsCommand();
        var settings = new ListActionsCommand.Settings();

        var context = new CommandContext(
            Array.Empty<string>(),
            new FakeRemainingArguments(),
            "actions",
            null);
        var output = CaptureOutput(() => command.Execute(context, settings, CancellationToken.None));
        using var document = JsonDocument.Parse(output);

        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        var commands = document.RootElement.GetProperty("commands");

        foreach (var (cliCommandName, _, _) in _CliCategoryMetadata.Categories)
        {
            Assert.True(commands.TryGetProperty(cliCommandName, out _), $"Missing command '{cliCommandName}'.");
        }
    }

    private static Type ResolveRegistryType(string cliCommandName)
    {
        var entry = _CliCategoryMetadata.Categories.First(c =>
            string.Equals(c.CliCommandName, cliCommandName, StringComparison.OrdinalIgnoreCase));

        // RegistryTypeName is emitted as "ServiceRegistry.Mail"; reflection needs "ServiceRegistry+Mail".
        var nestedName = entry.RegistryTypeName.Replace('.', '+');
        var type = typeof(_CliCategoryMetadata).Assembly.GetType($"OutlookMcp.Generated.{nestedName}");

        Assert.NotNull(type);
        return type!;
    }

    private static string[] GetExpectedActions(Type registryType)
    {
        var actionMethod = registryType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .First(m => m.Name == "ToActionString"
                && m.GetParameters().Length == 1
                && m.GetParameters()[0].ParameterType.IsEnum);

        var enumType = actionMethod.GetParameters()[0].ParameterType;
        var values = Enum.GetValues(enumType);
        var results = new List<string>(values.Length);

        foreach (var value in values)
        {
            var action = actionMethod.Invoke(null, [value]) as string;
            results.Add(action ?? string.Empty);
        }

        return results.OrderBy(action => action, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static string[] GetActualActions(Type registryType)
    {
        var validActionsField = registryType
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .First(f => f.Name == "ValidActions");

        var actions = (string[])validActionsField.GetValue(null)!;
        return actions.OrderBy(action => action, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static string CaptureOutput(Func<int> action)
    {
        var original = Console.Out;
        using var writer = new StringWriter();
        try
        {
            Console.SetOut(writer);
            action();
            return writer.ToString().Trim();
        }
        finally
        {
            Console.SetOut(original);
        }
    }

    private sealed class FakeRemainingArguments : IRemainingArguments
    {
        private static readonly ILookup<string, string?> EmptyLookup =
            Array.Empty<string>().ToLookup(_ => string.Empty, _ => (string?)null);

        public ILookup<string, string?> Parsed { get; } = EmptyLookup;
        public IReadOnlyList<string> Raw { get; } = Array.Empty<string>();
    }
}
