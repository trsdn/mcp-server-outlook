using OutlookMcp.Core.Commands.Folder;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Folder mutation: create, rename, move and delete (#15).
///
/// <para>
/// The <c>folder</c> tool could only read. Filing mail into a folder that does not exist yet had no
/// answer, so "archive these into a 2024 folder" ended at the first step.
/// </para>
///
/// <para>
/// <b>What makes this dangerous, and what these tests are really for.</b> Outlook's
/// <c>Folder.Delete</c> will happily delete the Inbox. There is no confirmation, the operation
/// succeeds, and everything filed in it goes with it. The same is true of rename and move. So the
/// guards - refusing a default folder and refusing a store root - are the substance of this feature,
/// not a nicety around it, and most of the tests below exercise a refusal rather than a success.
/// </para>
///
/// <para>
/// Everything created here lives under a GUID-named scratch folder in the default store and is
/// removed in <c>finally</c>. No pre-existing folder is ever renamed, moved or deleted: every
/// destructive assertion in this file is made against a folder this test created moments earlier.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "FolderMutation")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookFolderMutationTests(ITestOutputHelper output)
{
    /// <summary>
    /// The whole lifecycle end to end: create a child, confirm it is really there by listing the
    /// parent, then delete it and confirm it is gone.
    ///
    /// <para>
    /// The listing check is the point. A <c>create</c> that returned <c>success: true</c> without
    /// producing a folder would pass any assertion made only on its own return value, and that is
    /// exactly the failure this project keeps finding.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void Create_ThenDelete_ReallyAddsAndRemovesTheFolder()
    {
        var commands = new FolderCommands();
        string parent = ResolveScratchParent(commands);
        string name = ScratchName();

        string? created = null;
        try
        {
            var create = commands.Create(parent, name);
            Assert.True(create.Success, create.ErrorMessage);
            Assert.NotNull(create.FolderPath);
            created = create.FolderPath;

            output.WriteLine($"Created: {created}");

            Assert.Contains(
                commands.ListChildren(parent).Folders,
                f => string.Equals(f.Name, name, StringComparison.Ordinal));
        }
        finally
        {
            if (created != null)
            {
                var delete = commands.Delete(created);
                output.WriteLine($"Delete: success={delete.Success} {delete.ErrorMessage}");
                Assert.True(delete.Success, delete.ErrorMessage);
            }
        }

        Assert.DoesNotContain(
            commands.ListChildren(parent).Folders,
            f => string.Equals(f.Name, name, StringComparison.Ordinal));
    }

    /// <summary>
    /// Rename must change the name and leave the folder reachable at its new path. A rename that
    /// reported success but left the old name in place would be invisible to a caller who trusted
    /// the result.
    /// </summary>
    [SkippableFact]
    public void Rename_ChangesTheNameAndTheNewPathResolves()
    {
        var commands = new FolderCommands();
        string parent = ResolveScratchParent(commands);
        string original = ScratchName();
        string renamed = original + "-renamed";

        string? current = null;
        try
        {
            var create = commands.Create(parent, original);
            Assert.True(create.Success, create.ErrorMessage);
            current = create.FolderPath;

            var rename = commands.Rename(current, renamed);
            Assert.True(rename.Success, rename.ErrorMessage);
            Assert.Equal(renamed, rename.Name);
            Assert.NotNull(rename.FolderPath);
            current = rename.FolderPath;

            output.WriteLine($"Renamed to: {current}");

            // The path the rename reported must actually work, not merely look plausible. Outlook's
            // folder tree lags a rename, which is why the operation waits for it; if it gave up, it
            // says so in note rather than pretending the path is usable.
            Assert.Null(rename.Note);

            var resolved = commands.ResolvePath(current);
            Assert.True(resolved.Success, resolved.ErrorMessage);
            Assert.Equal(renamed, resolved.Name);

            Assert.DoesNotContain(
                commands.ListChildren(parent).Folders,
                f => string.Equals(f.Name, original, StringComparison.Ordinal));
        }
        finally
        {
            if (current != null)
            {
                SweepScratchFolders(commands, parent);
            }
        }
    }

    /// <summary>
    /// Move must reparent the folder: gone from the old parent, present under the new one, and
    /// reachable at the path the move reported.
    /// </summary>
    [SkippableFact]
    public void Move_ReparentsTheFolderAndTheNewPathResolves()
    {
        var commands = new FolderCommands();
        string parent = ResolveScratchParent(commands);
        string moverName = ScratchName();
        string hostName = ScratchName();

        string? host = null;
        string? mover = null;
        try
        {
            host = Assert.IsType<string>(commands.Create(parent, hostName).FolderPath);
            mover = Assert.IsType<string>(commands.Create(parent, moverName).FolderPath);

            var move = commands.Move(mover, host);
            Assert.True(move.Success, move.ErrorMessage);
            Assert.NotNull(move.FolderPath);
            mover = move.FolderPath;

            output.WriteLine($"Moved to: {mover}");

            Assert.Contains(
                commands.ListChildren(host).Folders,
                f => string.Equals(f.Name, moverName, StringComparison.Ordinal));
            Assert.DoesNotContain(
                commands.ListChildren(parent).Folders,
                f => string.Equals(f.Name, moverName, StringComparison.Ordinal));

            var resolved = commands.ResolvePath(mover);
            Assert.True(resolved.Success, resolved.ErrorMessage);
        }
        finally
        {
            // The host takes the moved child with it; the prefix sweep catches the mover too if the
            // move never happened and it is still a sibling.
            SweepScratchFolders(commands, host ?? parent);
            SweepScratchFolders(commands, parent);
        }
    }

    /// <summary>
    /// <b>The guard that matters most.</b> Outlook will delete the Inbox if asked. Nothing in COM
    /// refuses, nothing prompts, and every message in it goes too. So deleting a default folder must
    /// be refused here, by name, before it ever reaches Outlook.
    ///
    /// <para>
    /// The roles below are the tool's own aliases, and that detail is load-bearing: an earlier
    /// version of this test used <c>sentmail</c> and <c>deleteditems</c>, which are not aliases, so
    /// the refusal came from "could not be resolved" and the guard was never exercised at all. The
    /// test passed while proving nothing. Hence the message assertion.
    /// </para>
    /// </summary>
    [SkippableTheory]
    [InlineData("inbox")]
    [InlineData("sent")]
    [InlineData("deleted")]
    [InlineData("drafts")]
    [InlineData("calendar")]
    public void Delete_OfADefaultFolder_IsRefused(string role)
    {
        var commands = new FolderCommands();
        EnsureOutlookAvailable(commands);

        var result = commands.Delete(role);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);

        // It must be refused *because it is a default folder*, not because the alias was unknown.
        Assert.Contains(role, result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("could not be resolved", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Renaming or moving a default folder is the same hazard as deleting one - the folder survives,
    /// but every stored path pointing at it stops working, including Outlook's own. Refused for the
    /// same reason.
    /// </summary>
    [SkippableFact]
    public void RenameAndMove_OfADefaultFolder_AreRefused()
    {
        var commands = new FolderCommands();
        EnsureOutlookAvailable(commands);

        var rename = commands.Rename("inbox", "not-the-inbox");
        Assert.False(rename.Success);
        Assert.NotNull(rename.ErrorMessage);

        var move = commands.Move("inbox", "drafts");
        Assert.False(move.Success);
        Assert.NotNull(move.ErrorMessage);

        output.WriteLine($"Rename refused: {rename.ErrorMessage}");
        output.WriteLine($"Move refused: {move.ErrorMessage}");
    }

    /// <summary>
    /// A store root has no parent to be removed from, and deleting one would mean losing a whole
    /// mailbox. Refused explicitly rather than left to whatever COM does.
    /// </summary>
    [SkippableFact]
    public void Delete_OfAStoreRoot_IsRefused()
    {
        var commands = new FolderCommands();
        var stores = commands.ListStores();
        Skip.If(!stores.Success, stores.ErrorMessage);

        string? root = stores.Stores
            .Select(s => s.RootFolderPath)
            .FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
        Skip.If(root == null, "No store reported a root folder path.");

        var result = commands.Delete(root);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Creating a second child with a name already taken must fail with something a caller can act
    /// on, rather than a bare COM error - and must not silently return the existing folder, which
    /// would let a caller believe they had a fresh empty one.
    /// </summary>
    [SkippableFact]
    public void Create_WithANameAlreadyTaken_IsRefused()
    {
        var commands = new FolderCommands();
        string parent = ResolveScratchParent(commands);
        string name = ScratchName();

        string? created = null;
        try
        {
            created = Assert.IsType<string>(commands.Create(parent, name).FolderPath);

            var second = commands.Create(parent, name);

            Assert.False(second.Success);
            Assert.NotNull(second.ErrorMessage);
            Assert.Contains(name, second.ErrorMessage!, StringComparison.Ordinal);

            output.WriteLine($"Refused as expected: {second.ErrorMessage}");
        }
        finally
        {
            if (created != null)
            {
                SweepScratchFolders(commands, parent);
            }
        }
    }

    /// <summary>
    /// A blank name must be refused. Outlook accepts one in some builds and produces a folder that
    /// cannot be addressed by path afterwards.
    /// </summary>
    [SkippableFact]
    public void Create_WithABlankName_IsRefused()
    {
        var commands = new FolderCommands();
        string parent = ResolveScratchParent(commands);

        var result = commands.Create(parent, "   ");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Where scratch folders go: under the default store's Inbox, which every profile has and which
    /// is safe to add children to. Returns its path so the tests address it the same way a caller
    /// would.
    /// </summary>
    private static string ResolveScratchParent(FolderCommands commands)
    {
        EnsureOutlookAvailable(commands);

        var inbox = commands.ResolvePath("inbox");
        Skip.If(!inbox.Success, inbox.ErrorMessage);
        Skip.If(string.IsNullOrWhiteSpace(inbox.FolderPath), "Inbox reported no usable folder path.");

        return inbox.FolderPath!;
    }

    /// <summary>
    /// Removes every scratch folder left under the parent, by <b>name prefix</b> rather than by a
    /// path captured earlier.
    ///
    /// <para>
    /// That distinction is the whole point. An earlier version of these tests deleted by the path the
    /// operation had returned, and when the rename bug made that path stale the delete quietly failed
    /// - its return value was discarded - leaving four real folders in the developer's mailbox. They
    /// were found by listing the Inbox afterwards, not by any assertion. Sweeping by prefix cannot be
    /// defeated that way, and the failures are reported rather than swallowed.
    /// </para>
    /// </summary>
    private void SweepScratchFolders(FolderCommands commands, string parent)
    {
        // Two passes, because a single pass has already been observed to report clean while leaving
        // real folders behind. The second pass is what turns "nothing failed" into evidence.
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

        // Anything still here is a real folder left in a real mailbox. Say so loudly rather than
        // ending quietly, which is how the last four went unnoticed.
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

    private const string ScratchPrefix = "mcp-test-";

    private static string ScratchName() => $"{ScratchPrefix}{Guid.NewGuid():N}";

    private static void EnsureOutlookAvailable(FolderCommands commands)
    {
        var probe = commands.ListDefault();
        Skip.If(!probe.Success, probe.ErrorMessage ?? "Outlook is not available.");
    }
}
