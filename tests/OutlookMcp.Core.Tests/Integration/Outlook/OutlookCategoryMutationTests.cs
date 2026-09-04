using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Category creation, update and removal (#15).
///
/// <para>
/// <c>list-categories</c> could already discover the master list and <c>set-categories</c> could
/// write a name onto an item, but nothing could CREATE a category. Assigning a name that was not in
/// the list produced a colourless, unfilterable label. These tests exercise the write side against
/// the real per-mailbox master category list.
/// </para>
///
/// <para>
/// The master list is a shared user setting, so every category these tests touch is named with a
/// unique <c>mcp-test-cat-</c> GUID prefix and removed again in a <c>finally</c>. Nothing here ever
/// mutates a category it did not create.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailCategory")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookCategoryMutationTests(ITestOutputHelper output)
{
    private const string Prefix = "mcp-test-cat-";

    private static string ScratchName() => $"{Prefix}{Guid.NewGuid():N}";

    /// <summary>
    /// The core round-trip: a created category appears in <c>list-categories</c> carrying the exact
    /// friendly colour name it was created with. This is what makes a subsequent <c>set-categories</c>
    /// produce a real, filterable label rather than a colourless string.
    /// </summary>
    [SkippableFact]
    public void CreateCategory_WithAColour_IsListedWithThatColour()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string name = ScratchName();

        try
        {
            var created = commands.CreateCategory(name, color: "darkTeal");
            Assert.True(created.Success, created.ErrorMessage);
            Assert.Null(created.ErrorMessage);
            Assert.NotNull(created.Category);
            Assert.Equal(name, created.Category!.Name);
            Assert.Equal("darkTeal", created.Category.Color);

            var listed = commands.ListCategories();
            Assert.True(listed.Success, listed.ErrorMessage);

            var match = listed.Categories.FirstOrDefault(
                c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(match);
            Assert.Equal("darkTeal", match!.Color);
        }
        finally
        {
            Cleanup(commands, name);
        }
    }

    /// <summary>
    /// Creating a name that already exists is refused, not silently duplicated - a second entry of
    /// the same name would be unaddressable by <c>set-categories</c>.
    /// </summary>
    [SkippableFact]
    public void CreateCategory_ThatAlreadyExists_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string name = ScratchName();

        try
        {
            var first = commands.CreateCategory(name, color: "blue");
            Assert.True(first.Success, first.ErrorMessage);

            var second = commands.CreateCategory(name, color: "red");
            Assert.False(second.Success);
            Assert.NotNull(second.ErrorMessage);
            Assert.Contains("already exists", second.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Cleanup(commands, name);
        }
    }

    /// <summary>
    /// A blank name is refused before any COM call - the name is how the category is addressed
    /// everywhere else, so an empty one is meaningless.
    /// </summary>
    [SkippableFact]
    public void CreateCategory_WithABlankName_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().CreateCategory("   ");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    /// <summary>
    /// An unrecognised colour is not a failure: the category is still created, but with no colour,
    /// and the result says so explicitly rather than pretending a colour was applied. This mirrors
    /// how Outlook itself tolerates the absence of a colour.
    /// </summary>
    [SkippableFact]
    public void CreateCategory_WithAnUnknownColour_SucceedsWithNoColourAndSaysSo()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string name = ScratchName();

        try
        {
            var created = commands.CreateCategory(name, color: "chartreuse-ish");
            Assert.True(created.Success, created.ErrorMessage);
            Assert.Null(created.ErrorMessage);
            Assert.NotNull(created.Category);
            Assert.Equal("none", created.Category!.Color);
            Assert.NotNull(created.Message);
            Assert.Contains("not recognised", created.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Cleanup(commands, name);
        }
    }

    /// <summary>
    /// A misspelled shortcut IS refused on create. Unlike a colour, a shortcut has no visible "none"
    /// the user would notice was missing, so a silent downgrade would hide the mistake.
    /// </summary>
    [SkippableFact]
    public void CreateCategory_WithAnUnknownShortcut_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string name = ScratchName();
        bool created = false;

        try
        {
            var result = commands.CreateCategory(name, color: "blue", shortcutKey: "ctrlF99");
            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Contains("shortcut", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);

            // Prove the refusal was total: nothing was created as a side effect.
            var listed = commands.ListCategories();
            created = listed.Categories.Any(
                c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            Assert.False(created, "A refused create must not leave a category behind.");
        }
        finally
        {
            if (created)
            {
                Cleanup(commands, name);
            }
        }
    }

    /// <summary>
    /// Update changes colour and name together, and the change is visible on a subsequent read of
    /// the master list under the new name.
    /// </summary>
    [SkippableFact]
    public void UpdateCategory_RecolourAndRename_IsReflectedInTheList()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string original = ScratchName();
        string renamed = ScratchName();
        string cleanupName = original;

        try
        {
            var created = commands.CreateCategory(original, color: "blue");
            Assert.True(created.Success, created.ErrorMessage);

            var updated = commands.UpdateCategory(original, newName: renamed, color: "purple");
            Assert.True(updated.Success, updated.ErrorMessage);
            Assert.Null(updated.ErrorMessage);
            cleanupName = renamed;

            var listed = commands.ListCategories();
            Assert.True(listed.Success, listed.ErrorMessage);

            Assert.DoesNotContain(
                listed.Categories,
                c => string.Equals(c.Name, original, StringComparison.OrdinalIgnoreCase));

            var match = listed.Categories.FirstOrDefault(
                c => string.Equals(c.Name, renamed, StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(match);
            Assert.Equal("purple", match!.Color);
        }
        finally
        {
            Cleanup(commands, cleanupName);
        }
    }

    /// <summary>
    /// Updating a name that is not in the list is refused, so a caller cannot mistake a typo for a
    /// successful edit.
    /// </summary>
    [SkippableFact]
    public void UpdateCategory_ThatDoesNotExist_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().UpdateCategory(ScratchName(), color: "red");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("no category", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Delete removes the category, and a subsequent list confirms it is gone.
    /// </summary>
    [SkippableFact]
    public void DeleteCategory_RemovesItFromTheList()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string name = ScratchName();
        bool needsCleanup = true;

        try
        {
            var created = commands.CreateCategory(name, color: "green");
            Assert.True(created.Success, created.ErrorMessage);

            var deleted = commands.DeleteCategory(name);
            Assert.True(deleted.Success, deleted.ErrorMessage);
            Assert.Null(deleted.ErrorMessage);
            needsCleanup = false;

            var listed = commands.ListCategories();
            Assert.True(listed.Success, listed.ErrorMessage);
            Assert.DoesNotContain(
                listed.Categories,
                c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            if (needsCleanup)
            {
                Cleanup(commands, name);
            }
        }
    }

    /// <summary>
    /// Deleting a name that is not present is refused rather than reported as a no-op success, so a
    /// caller can distinguish a real removal from a name that was never there.
    /// </summary>
    [SkippableFact]
    public void DeleteCategory_ThatDoesNotExist_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().DeleteCategory(ScratchName());

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("no category", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    private void Cleanup(MailCommands commands, string name)
    {
        try
        {
            _ = commands.DeleteCategory(name);
        }
        catch (Exception ex)
        {
            output.WriteLine($"Cleanup failed for category '{name}': {ex.Message}");
        }
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
