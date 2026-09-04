using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.Mail;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Folder empty (#15) - deleting a folder's contents while keeping the folder.
///
/// <para>
/// <b>Why this file is mostly refusals.</b> Emptying the Inbox in one call is unrecoverable in a way
/// deleting one message is not, so the guards ARE the feature: default and special folders and store
/// roots are refused outright, and even an ordinary folder is refused without <c>confirm=true</c>.
/// The happy path - emptying a scratch folder - is the least important test here, so it is one test
/// among many that prove the refusals.
/// </para>
///
/// <para>
/// Every folder emptied here is a GUID-named scratch folder created moments earlier under the
/// default Inbox and swept away in <c>finally</c>. No pre-existing folder is ever emptied. The one
/// item these tests put in a scratch folder is a draft they created; emptying moves it to Deleted
/// Items, which is the documented behaviour.
/// </para>
///
/// <para>
/// <b>Chosen semantics (asserted below):</b> empty clears the folder's own items only, moving each
/// to Deleted Items; subfolders and their contents are left untouched.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "FolderEmpty")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookFolderEmptyTests(ITestOutputHelper output)
{
    private const string ScratchPrefix = "mcp-test-";

    private static string ScratchName() => $"{ScratchPrefix}{Guid.NewGuid():N}";

    /// <summary>
    /// The most important refusal: emptying a default or special folder is rejected before anything
    /// is touched, and rejected BECAUSE it is that folder - not because the alias failed to resolve.
    /// </summary>
    [SkippableTheory]
    [InlineData("inbox")]
    [InlineData("sent")]
    [InlineData("drafts")]
    [InlineData("deleted")]
    [InlineData("calendar")]
    public void Empty_OfADefaultFolder_IsRefused(string role)
    {
        var commands = new FolderCommands();
        EnsureOutlookAvailable(commands);

        // confirm=true so the refusal cannot be the confirmation gate - it must be the guard.
        var result = commands.Empty(role, confirm: true);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains(role, result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("could not be resolved", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// A store root has no business being emptied wholesale; refused like a default folder.
    /// </summary>
    [SkippableFact]
    public void Empty_OfAStoreRoot_IsRefused()
    {
        var commands = new FolderCommands();
        var stores = commands.ListStores();
        Skip.If(!stores.Success, stores.ErrorMessage);

        string? root = stores.Stores
            .Select(s => s.RootFolderPath)
            .FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
        Skip.If(root == null, "No store reported a root folder path.");

        var result = commands.Empty(root, confirm: true);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.DoesNotContain("could not be resolved", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Blank target is refused before any COM call - empty has no default folder.
    /// </summary>
    [SkippableFact]
    public void Empty_WithABlankFolder_IsRefused()
    {
        var result = new FolderCommands().Empty("   ", confirm: true);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    /// <summary>
    /// Without <c>confirm=true</c> the call is refused AND nothing is removed. The item put in the
    /// scratch folder is still there afterwards - the refusal is proven by the surviving item, not
    /// only by the return value.
    /// </summary>
    [SkippableFact]
    public void Empty_WithoutConfirmation_IsRefusedAndRemovesNothing()
    {
        var folders = new FolderCommands();
        var mail = new MailCommands();
        string parent = ResolveScratchParent(folders);

        try
        {
            string scratchPath = CreateScratchFolderWithOneItem(folders, mail, parent);

            var result = folders.Empty(scratchPath, confirm: false);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Contains("confirm", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            // The refusal must not have deleted anything.
            var after = folders.ResolvePath(scratchPath);
            Assert.True(after.Success, after.ErrorMessage);
            Assert.Equal(1, after.ItemCount);
        }
        finally
        {
            SweepScratchFolders(folders, parent);
        }
    }

    /// <summary>
    /// The happy path: with confirmation, the folder's item is removed, the count of removed items is
    /// reported, and the folder itself survives with zero items.
    /// </summary>
    [SkippableFact]
    public void Empty_WithConfirmation_RemovesItemsButKeepsTheFolder()
    {
        var folders = new FolderCommands();
        var mail = new MailCommands();
        string parent = ResolveScratchParent(folders);

        try
        {
            string scratchPath = CreateScratchFolderWithOneItem(folders, mail, parent);

            var result = folders.Empty(scratchPath, confirm: true);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(1, result.ItemsRemoved);

            // The folder is still there, now empty.
            var after = folders.ResolvePath(scratchPath);
            Assert.True(after.Success, after.ErrorMessage);
            Assert.True(after.Resolved);
            Assert.Equal(0, after.ItemCount);
        }
        finally
        {
            SweepScratchFolders(folders, parent);
        }
    }

    /// <summary>
    /// An already-empty folder empties to a success reporting zero removed - distinguishable from a
    /// refusal, which returns <c>Success = false</c>.
    /// </summary>
    [SkippableFact]
    public void Empty_OfAnEmptyFolder_SucceedsWithZeroRemoved()
    {
        var folders = new FolderCommands();
        string parent = ResolveScratchParent(folders);
        string name = ScratchName();

        try
        {
            var created = folders.Create(parent, name);
            Assert.True(created.Success, created.ErrorMessage);

            var result = folders.Empty(created.FolderPath, confirm: true);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(0, result.ItemsRemoved);
        }
        finally
        {
            SweepScratchFolders(folders, parent);
        }
    }

    /// <summary>
    /// Empty clears the folder's own items but leaves a subfolder in place. This nails down the one
    /// genuine design decision - "empty the archive" does not recurse into its subfolders.
    /// </summary>
    [SkippableFact]
    public void Empty_LeavesSubfoldersUntouched()
    {
        var folders = new FolderCommands();
        var mail = new MailCommands();
        string parent = ResolveScratchParent(folders);
        string? scratchPath = null;

        try
        {
            scratchPath = CreateScratchFolderWithOneItem(folders, mail, parent);

            string childName = ScratchName();
            var child = folders.Create(scratchPath, childName);
            Assert.True(child.Success, child.ErrorMessage);

            var result = folders.Empty(scratchPath, confirm: true);
            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(1, result.ItemsRemoved);

            // The subfolder survived: it is still a child of the emptied folder.
            var children = folders.ListChildren(scratchPath);
            Assert.True(children.Success, children.ErrorMessage);
            Assert.Contains(
                children.Folders,
                f => string.Equals(f.Name, childName, StringComparison.Ordinal));
        }
        finally
        {
            SweepScratchFolders(folders, parent);
        }
    }

    /// <summary>
    /// Creates a GUID-named scratch folder under the parent and puts exactly one item in it (a draft,
    /// moved in from Drafts). Returns the scratch folder's path.
    /// </summary>
    private static string CreateScratchFolderWithOneItem(
        FolderCommands folders, MailCommands mail, string parent)
    {
        string name = ScratchName();
        var created = folders.Create(parent, name);
        Assert.True(created.Success, created.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(created.FolderPath), "Scratch folder reported no path.");

        var draft = mail.CreateMailDraft(
            recipientTo: "nobody@example.invalid",
            subject: $"{ScratchPrefix}{Guid.NewGuid():N}",
            body: "Folder empty test item.");
        Assert.True(draft.Success, draft.ErrorMessage);

        var moved = mail.Move(targetFolder: created.FolderPath!, entryId: draft.EntryId, useActiveMail: false);
        Assert.True(moved.Success, moved.ErrorMessage);

        return created.FolderPath!;
    }

    private static string ResolveScratchParent(FolderCommands commands)
    {
        EnsureOutlookAvailable(commands);

        var inbox = commands.ResolvePath("inbox");
        Skip.If(!inbox.Success, inbox.ErrorMessage);
        Skip.If(string.IsNullOrWhiteSpace(inbox.FolderPath), "Inbox reported no usable folder path.");

        return inbox.FolderPath!;
    }

    /// <summary>
    /// Removes every scratch folder left under the parent by name prefix, in two passes, reporting
    /// loudly if anything real is left behind.
    /// </summary>
    private void SweepScratchFolders(FolderCommands commands, string parent)
    {
        for (int pass = 1; pass <= 2; pass++)
        {
            var listing = commands.ListChildren(parent);
            if (!listing.Success)
            {
                output.WriteLine($"Sweep could not list '{parent}': {listing.ErrorMessage}");
                return;
            }

            var leftovers = listing.Folders
                .Where(f => f.Name?.StartsWith(ScratchPrefix, StringComparison.Ordinal) == true)
                .ToList();

            if (leftovers.Count == 0)
            {
                return;
            }

            foreach (var folder in leftovers)
            {
                var deleted = commands.Delete(folder.FolderPath);
                output.WriteLine($"Sweep pass {pass}: {folder.Name} -> success={deleted.Success} {deleted.ErrorMessage}");
            }
        }

        var remaining = commands.ListChildren(parent).Folders
            .Where(f => f.Name?.StartsWith(ScratchPrefix, StringComparison.Ordinal) == true)
            .Select(f => f.Name)
            .ToList();

        if (remaining.Count > 0)
        {
            output.WriteLine(
                $"SWEEP FAILED - {remaining.Count} scratch folder(s) remain in '{parent}': "
                + string.Join(", ", remaining));
        }
    }

    private static void EnsureOutlookAvailable(FolderCommands commands)
    {
        var probe = commands.ListDefault();
        Skip.If(!probe.Success, probe.ErrorMessage ?? "Outlook is not available.");
    }
}
