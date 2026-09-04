using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Store and account discovery, and per-store default folders (#38).
///
/// <para>
/// Everything in this surface used to target the default delivery store implicitly.
/// <c>folder.list-default</c> reported one account's folders and said nothing about the others
/// existing at all, which is the failure mode this project keeps running into: a confident answer to
/// a question the caller did not ask. A profile with an Exchange mailbox plus an archive or a PST
/// would be told, with <c>success: true</c>, that its Inbox was the only Inbox.
/// </para>
///
/// <para>
/// These tests are deliberately written not to assume the owner's profile has more than one store.
/// A profile with exactly one store is a legitimate configuration, and a test that only passes on a
/// multi-store profile would be a test that silently stops testing anything. What is asserted
/// instead is the property that must hold either way: whatever stores exist are all enumerated, they
/// are addressable by id, and every folder result names the store it came from.
/// </para>
///
/// <para>
/// Everything here is read-only. Nothing is created, moved, sent or deleted.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Store")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookStoreTests(ITestOutputHelper output)
{
    /// <summary>
    /// The baseline: a profile always has at least one store, and every store must arrive with the
    /// id needed to address it. A display name alone is not enough - two stores can share one.
    /// </summary>
    [SkippableFact]
    public void ListStores_ReturnsEveryStoreWithAnAddressableId()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().ListStores();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Stores);

        foreach (var store in result.Stores)
        {
            Assert.False(string.IsNullOrWhiteSpace(store.StoreId), "A store arrived without an id.");
            Assert.False(string.IsNullOrWhiteSpace(store.DisplayName), "A store arrived without a name.");
            output.WriteLine(
                $"{store.DisplayName} | default={store.IsDefaultStore} | dataFile={store.IsDataFileStore} "
                + $"| type={store.ExchangeStoreType} | account={store.AccountSmtpAddress ?? "(none)"}");
        }

        output.WriteLine($"{result.Stores.Count} store(s) in this profile.");
    }

    /// <summary>
    /// Exactly one store is the default delivery store. Zero would mean the tool cannot tell a caller
    /// which mailbox an unqualified request lands in; more than one would mean it is guessing.
    /// </summary>
    [SkippableFact]
    public void ListStores_MarksExactlyOneStoreAsTheDefault()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().ListStores();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Stores, store => store.IsDefaultStore);
    }

    /// <summary>
    /// The reason this issue exists: a folder listing that does not say which mailbox it read is
    /// unusable on a multi-store profile, because every store has an Inbox.
    /// </summary>
    [SkippableFact]
    public void ListDefault_NamesTheStoreEachFolderCameFrom()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().ListDefault();

        Assert.True(result.Success, result.ErrorMessage);

        var available = result.Folders.Where(f => f.Available).ToList();
        Assert.NotEmpty(available);

        foreach (var folder in available)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(folder.StoreId),
                $"Folder role '{folder.Role}' arrived without a store id.");
            Assert.False(
                string.IsNullOrWhiteSpace(folder.StoreName),
                $"Folder role '{folder.Role}' arrived without a store name.");
        }

        output.WriteLine($"{available.Count} default folder(s), all naming their store.");
    }

    /// <summary>
    /// <c>available: true</c> has to mean the caller can actually reach the folder, and the only
    /// handle this surface offers is the path. This is the assertion that caught the real bug:
    /// <c>Store.GetDefaultFolder</c> answers for roles a store does not have, returning a folder
    /// object that is not in the tree and whose <c>FolderPath</c> degenerates to its entry id - a
    /// long hex string that looks like a value and resolves to nothing. On the developer's archive
    /// that was nine roles out of ten, every one of them reported as available.
    /// </summary>
    [SkippableFact]
    public void ListDefault_ReportsARoleAvailableOnlyWhenItsPathCanBeUsed()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();

        var stores = commands.ListStores();
        Assert.True(stores.Success, stores.ErrorMessage);

        foreach (var store in stores.Stores)
        {
            var folders = commands.ListDefault(storeId: store.StoreId);
            Assert.True(folders.Success, folders.ErrorMessage);

            foreach (var folder in folders.Folders.Where(f => f.Available))
            {
                Assert.NotNull(folder.FolderPath);
                Assert.StartsWith(
                    @"\\",
                    folder.FolderPath!,
                    StringComparison.Ordinal);
            }

            foreach (var folder in folders.Folders.Where(f => !f.Available))
            {
                // An unavailable role must not carry a path either - a caller filtering on
                // folderPath rather than available would otherwise still be misled.
                Assert.Null(folder.FolderPath);
            }

            output.WriteLine(
                $"{store.DisplayName}: available roles {string.Join(", ", folders.Folders.Where(f => f.Available).Select(f => f.Role))}");
        }
    }

    /// <summary>
    /// Every path this surface hands out must be resolvable by it. A value that cannot be fed back
    /// in is worse than no value, because the caller only finds out one call later and cannot tell
    /// the difference between "wrong path" and "folder is empty".
    /// </summary>
    [SkippableFact]
    public void ListDefault_PathsForEveryStore_ResolveBack()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();

        var stores = commands.ListStores();
        Assert.True(stores.Success, stores.ErrorMessage);

        int checkedPaths = 0;

        foreach (var store in stores.Stores)
        {
            var folders = commands.ListDefault(storeId: store.StoreId);
            Assert.True(folders.Success, folders.ErrorMessage);

            foreach (var folder in folders.Folders.Where(f => f.Available))
            {
                var resolved = commands.ResolvePath(folder.FolderPath, includeItemCount: false);

                Assert.True(
                    resolved.Success && resolved.Resolved,
                    $"'{folder.FolderPath}' was reported by list-default but resolve-path could not "
                    + $"find it: {resolved.ErrorMessage}");

                checkedPaths++;
            }
        }

        Assert.True(checkedPaths > 0, "No folder paths were checked, so this test proved nothing.");
        output.WriteLine($"{checkedPaths} folder path(s) round-tripped through resolve-path.");
    }

    /// <summary>
    /// Targeting the default store explicitly must return the same folders as not targeting anything.
    /// If these disagree, one of the two paths is reading a mailbox the caller did not ask for.
    /// </summary>
    [SkippableFact]
    public void ListDefault_TargetedAtTheDefaultStore_MatchesTheUntargetedCall()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();

        var stores = commands.ListStores();
        Assert.True(stores.Success, stores.ErrorMessage);

        var defaultStore = stores.Stores.Single(store => store.IsDefaultStore);

        var untargeted = commands.ListDefault();
        var targeted = commands.ListDefault(storeId: defaultStore.StoreId);

        Assert.True(targeted.Success, targeted.ErrorMessage);

        foreach (var role in untargeted.Folders.Where(f => f.Available).Select(f => f.Role))
        {
            var match = targeted.Folders.SingleOrDefault(f => f.Role == role);
            Assert.NotNull(match);
            Assert.True(match!.Available, $"Role '{role}' was available untargeted but not when targeted.");
            Assert.Equal(
                untargeted.Folders.Single(f => f.Role == role).FolderPath,
                match.FolderPath);
        }

        output.WriteLine($"Targeted and untargeted listings agree for store '{defaultStore.DisplayName}'.");
    }

    /// <summary>
    /// An unknown store id must fail, not quietly fall back to the default mailbox. Silently reading
    /// a different mailbox than the one requested is the worst available outcome here: the caller
    /// gets real folders, real item counts and <c>success: true</c>, from the wrong account.
    /// </summary>
    [SkippableFact]
    public void ListDefault_WithAnUnknownStoreId_IsRefusedRatherThanFallingBack()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().ListDefault(storeId: "0000NOTASTORE0000");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("store", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(result.Folders);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Every store id handed out by discovery must be usable for targeting. A profile where one of
    /// the enumerated stores cannot be opened means discovery is advertising something unreachable,
    /// which is worse than not listing it - so if a store legitimately cannot serve default folders
    /// it must say so through <c>available: false</c>, not through a failed call.
    /// </summary>
    [SkippableFact]
    public void ListDefault_IsAcceptedForEveryStoreDiscoveryReturns()
    {
        EnsureOutlookAvailable();

        var commands = new FolderCommands();

        var stores = commands.ListStores();
        Assert.True(stores.Success, stores.ErrorMessage);

        foreach (var store in stores.Stores)
        {
            var folders = commands.ListDefault(storeId: store.StoreId);

            Assert.True(
                folders.Success,
                $"Store '{store.DisplayName}' was enumerated but could not be targeted: {folders.ErrorMessage}");

            // A role a store does not have must explain itself. "available: false" with nothing else
            // reads as a transient failure; the caller needs to know the folder is simply not there
            // and where to look instead.
            foreach (var folder in folders.Folders.Where(f => !f.Available))
            {
                Assert.False(
                    string.IsNullOrWhiteSpace(folder.Note),
                    $"Role '{folder.Role}' in '{store.DisplayName}' is unavailable with no explanation.");
            }

            int available = folders.Folders.Count(f => f.Available);
            output.WriteLine($"{store.DisplayName}: {available} of {folders.Folders.Count} default role(s) available.");
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
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
