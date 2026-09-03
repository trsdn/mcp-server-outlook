using System.Text.Json;
using OutlookMcp.ComInterop.ServiceClient;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Category discovery (#15).
///
/// <para>
/// <c>mail set-categories</c> already shipped, but it writes a raw string into a field Outlook does
/// not validate. Assigning a category that is not in the master list succeeds, returns
/// <c>success: true</c>, and produces a category the user cannot filter or colour by - the exact
/// shape of failure this project keeps finding, where a check reports success without having
/// verified anything. Without discovery, an agent has no way to write a category that is real, so
/// it can only guess.
/// </para>
///
/// <para>
/// Everything here is read-only. The master category list is a user setting shared across every
/// item in the mailbox, so these tests never add, rename or remove an entry in it.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailCategory")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookCategoryTests(ITestOutputHelper output)
{
    /// <summary>
    /// The baseline: the master list is enumerated at all, and every entry carries the name that
    /// <c>set-categories</c> expects. A listing without names would be unusable for the one purpose
    /// it exists to serve.
    /// </summary>
    [SkippableFact]
    public void ListCategories_ReturnsTheMasterCategoryListWithUsableNames()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListCategories();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Categories);

        foreach (var category in result.Categories)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(category.Name),
                "A category arrived without a name, so it cannot be passed to set-categories.");
        }

        output.WriteLine($"{result.Categories.Count} categories: "
            + string.Join(", ", result.Categories.Select(c => $"{c.Name} ({c.Color})")));
    }

    /// <summary>
    /// Colour must arrive as a name, not as Outlook's raw enum ordinal. "4" tells a model nothing and
    /// cannot be repeated back to the user; "yellow" is the thing they actually see in Outlook.
    /// </summary>
    [SkippableFact]
    public void ListCategories_ReportsColourByNameRatherThanEnumOrdinal()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListCategories();

        Assert.True(result.Success, result.ErrorMessage);

        foreach (var category in result.Categories)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(category.Color),
                $"Category '{category.Name}' reported no colour at all.");

            Assert.False(
                int.TryParse(category.Color, out _),
                $"Category '{category.Name}' reported colour '{category.Color}', a raw enum ordinal. "
                + "A number is not something a model can show the user or reason about.");
        }
    }

    /// <summary>
    /// Names are what <c>set-categories</c> takes, so a duplicate name would make a category
    /// unaddressable: the caller would have no way to say which of the two they meant. Outlook
    /// enforces this, and this test exists to catch a listing that invents duplicates by, say,
    /// enumerating the collection twice.
    /// </summary>
    [SkippableFact]
    public void ListCategories_ReturnsEachCategoryExactlyOnce()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListCategories();

        Assert.True(result.Success, result.ErrorMessage);

        var duplicates = result.Categories
            .GroupBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();

        Assert.True(
            duplicates.Count == 0,
            $"These names appeared more than once: {string.Join(", ", duplicates)}");
    }

    /// <summary>
    /// The point of discovery is that its output feeds straight back into <c>set-categories</c>.
    /// This closes that loop against a real mailbox rather than assuming the two agree: the category
    /// is applied to a draft these tests created, read back, and the draft is deleted again.
    /// </summary>
    [SkippableFact]
    public void ListCategories_NamesAreAcceptedBySetCategories()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();

        var categories = commands.ListCategories();
        Assert.True(categories.Success, categories.ErrorMessage);
        Skip.If(categories.Categories.Count == 0, "This profile has no categories defined.");

        string name = categories.Categories[0].Name;
        string? draftId = null;

        try
        {
            var draft = commands.CreateMailDraft(
                recipientTo: "nobody@example.invalid",
                subject: $"{Marker} {Guid.NewGuid():N}",
                body: "Category round-trip check.");

            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            var applied = commands.SetCategories(categories: name, entryId: draftId, useActiveMail: false);
            Assert.True(applied.Success, applied.ErrorMessage);

            var read = commands.Read(entryId: draftId, useActiveMail: false);
            Assert.True(read.Success, read.ErrorMessage);

            Assert.Contains(name, read.Categories);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(draftId))
            {
                try
                {
                    _ = commands.Delete(entryId: draftId, useActiveMail: false);
                }
                catch (Exception ex)
                {
                    output.WriteLine($"Cleanup failed for {draftId}: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// The wire shape is what a model actually sees. Asserting on the object would pass even if the
    /// field never left the process - a vacuous assertion this project has already been bitten by.
    /// </summary>
    [SkippableFact]
    public void ListCategories_SerialisesColourAndNameOntoTheWire()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListCategories();
        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Categories.Count == 0, "This profile has no categories defined.");

        string json = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions);

        Assert.Contains("\"categories\"", json, StringComparison.Ordinal);
        Assert.Contains("\"color\"", json, StringComparison.Ordinal);
        Assert.Contains($"\"{result.Categories[0].Name}\"", json, StringComparison.Ordinal);
    }

    private const string Marker = "OutlookMcp category test";

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
