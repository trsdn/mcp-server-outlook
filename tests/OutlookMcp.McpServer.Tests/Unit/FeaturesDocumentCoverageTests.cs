using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Commands.Mail;
using Xunit;

namespace OutlookMcp.McpServer.Tests.Unit;

/// <summary>
/// Pins <c>FEATURES.md</c> to the generated tool surface, so a stale or silently truncated
/// operation count fails the build instead of shipping.
/// </summary>
/// <remarks>
/// <para>
/// <c>FEATURES.md</c> is hand-written prose describing a surface that is generated from
/// <c>[ServiceCategory]</c> / <c>[ServiceAction]</c>. Nothing read it: no test parsed it and the
/// build does not consume it, so it could disagree with the code indefinitely.
/// </para>
/// <para>
/// That mattered because the headline count is repeated across more than ten files, which makes it
/// the single most conflict-prone artefact in the repository. Every branch that adds an action
/// rewrites the same lines, so any two concurrent branches collide there. Resolving such a conflict
/// by taking one side is last-writer-wins: the losing side's sections are deleted and nothing
/// reports it. Reviewing two pull requests that each merge cleanly is not enough, because the loss
/// only appears once both have landed.
/// </para>
/// <para>
/// These assertions are deliberately name-independent. Section headings are English prose
/// ("Out-of-office / automatic replies") and do not map mechanically onto category ids, so matching
/// on names would need a hand-maintained alias table - the very thing that went stale in #138.
/// Section count, subtotal sum and headline together catch a dropped section, a stale per-tool
/// subtotal and a stale headline without one.
/// </para>
/// <para>
/// Permitted under the ADR-001 narrow exception: no COM object is touched, the subject is a parse
/// plus reflection over generated metadata, and the assertions fail when the document is wrong.
/// </para>
/// </remarks>
public class FeaturesDocumentCoverageTests
{
    private static readonly Regex SectionPattern = new(
        @"^##\s+(?<name>.+?)\s+Operations?\s+\((?<count>\d+)\s+operations?\)\s*$",
        RegexOptions.Multiline | RegexOptions.CultureInvariant);

    private static readonly Regex HeadlinePattern = new(
        @"(?<tools>\d+)\s+tools?\s+(?:with|and)\s+(?<operations>\d+)\s+operations",
        RegexOptions.CultureInvariant);

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "Documentation")]
    public void FeaturesDocument_IsFound_AndDeclaresSections()
    {
        // Guards against the locator or the parse silently finding nothing, which would make every
        // assertion below vacuously true - the failure mode this whole suite exists to prevent.
        var sections = ParseSections(ReadFeaturesDocument());

        Assert.NotEmpty(sections);
        Assert.NotEmpty(GetServiceCategories());
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "Documentation")]
    public void Features_DeclaresOneSectionPerGeneratedTool()
    {
        var sections = ParseSections(ReadFeaturesDocument());
        var categories = GetServiceCategories();

        Assert.True(sections.Count == categories.Count,
            $"FEATURES.md declares {sections.Count} operation sections but the generated surface has " +
            $"{categories.Count} tools ({string.Join(", ", categories.Keys.OrderBy(k => k, StringComparer.Ordinal))}). " +
            $"Sections present: {string.Join(", ", sections.Select(s => s.Name))}. " +
            "A missing section usually means a documentation merge conflict was resolved by taking " +
            "one side, which deletes the other side's sections.");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "Documentation")]
    public void Features_SubtotalsSumToTheGeneratedOperationCount()
    {
        var sections = ParseSections(ReadFeaturesDocument());
        var declared = sections.Sum(s => s.Count);
        var actual = GetTotalActionCount();

        Assert.True(declared == actual,
            $"FEATURES.md subtotals sum to {declared} but the generated surface has {actual} operations. " +
            $"Per-section: {string.Join(", ", sections.Select(s => $"{s.Name}={s.Count}"))}.");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "Documentation")]
    public void Features_HeadlineMatchesTheGeneratedSurface()
    {
        var text = ReadFeaturesDocument();
        var match = HeadlinePattern.Match(text);

        Assert.True(match.Success,
            "FEATURES.md has no 'N tools with M operations' headline. The headline is the figure " +
            "copied into every README, so it must be present and checkable.");

        var tools = int.Parse(match.Groups["tools"].Value, CultureInfo.InvariantCulture);
        var operations = int.Parse(match.Groups["operations"].Value, CultureInfo.InvariantCulture);

        Assert.True(tools == GetServiceCategories().Count && operations == GetTotalActionCount(),
            $"FEATURES.md headline claims {tools} tools / {operations} operations but the generated " +
            $"surface is {GetServiceCategories().Count} tools / {GetTotalActionCount()} operations.");
    }

    // ── Helpers ──────────────────────────────────────────────

    private sealed record Section(string Name, int Count);

    private static List<Section> ParseSections(string text)
    {
        return SectionPattern.Matches(text)
            .Select(m => new Section(
                m.Groups["name"].Value,
                int.Parse(m.Groups["count"].Value, CultureInfo.InvariantCulture)))
            .ToList();
    }

    private static string ReadFeaturesDocument()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "FEATURES.md");
            if (File.Exists(candidate))
            {
                return File.ReadAllText(candidate);
            }

            directory = directory.Parent;
        }

        throw new FileNotFoundException(
            $"FEATURES.md not found walking up from '{AppContext.BaseDirectory}'.");
    }

    private static Dictionary<string, Type> GetServiceCategories()
    {
        return typeof(IMailCommands).Assembly
            .GetTypes()
            .Where(t => t.IsInterface)
            .Select(t => (Type: t, Attr: t.GetCustomAttribute<ServiceCategoryAttribute>()))
            .Where(x => x.Attr is not null)
            .ToDictionary(x => x.Attr!.Category, x => x.Type, StringComparer.Ordinal);
    }

    private static int GetTotalActionCount()
    {
        return GetServiceCategories().Values.Sum(CountServiceActions);
    }

    private static int CountServiceActions(Type interfaceType)
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
