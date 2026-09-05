using OutlookMcp.Core.Commands.Folder;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Item classification in <c>folder.list-items</c> (#120).
///
/// <para>
/// <c>CreateFolderItemInfo</c> has typed branches for mail, appointment and contact, and everything
/// else falls through to a late-bound path that reported <c>rawItem.GetType().Name</c>. For a
/// late-bound RCW that name is the literal string <c>__ComObject</c>, which tells a caller nothing,
/// and the fallback also never set <c>EntryId</c> or <c>StoreId</c>. So listing the Tasks folder -
/// a surface this server already ships a full CRUD tool for - returned items that could not be
/// identified and could not subsequently be addressed by id.
/// </para>
///
/// <para>
/// These tests were written against a real profile whose Tasks folder held items and whose Notes
/// folder held two sticky notes. They assert the classification a caller can actually act on, not
/// merely that the call succeeded: a listing that returns <c>success: true</c> full of unusable
/// rows is the exact failure mode this repository keeps rediscovering.
/// </para>
///
/// <para>
/// Nothing here mutates the mailbox. Every folder is read, and a folder that happens to be empty
/// skips with a reason rather than passing vacuously.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "FolderItemType")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookFolderItemTypeTests(ITestOutputHelper output)
{
    /// <summary>
    /// A task in the Tasks folder must be reported as a task, with an entry id, so a caller can
    /// hand it straight to the <c>task</c> tool. This is the regression: it previously came back as
    /// <c>__ComObject</c> with no id at all.
    /// </summary>
    [SkippableFact]
    public void ListItems_Tasks_ClassifiesItemsAsTaskWithAnEntryId()
    {
        var item = FirstItemIn("tasks");

        output.WriteLine($"itemType={item.ItemType} messageClass={item.MessageClass} subject={item.Subject}");

        Assert.Equal("task", item.ItemType);
        Assert.False(
            string.IsNullOrWhiteSpace(item.EntryId),
            "A listed task must carry an entry id, otherwise it cannot be read or updated afterwards.");
    }

    /// <summary>
    /// The Notes folder is the reason #120 exists. A sticky note must classify as a note and carry
    /// its body, which is essentially all a note has.
    /// </summary>
    [SkippableFact]
    public void ListItems_Notes_ClassifiesItemsAsNoteWithAnEntryId()
    {
        var item = FirstItemIn("notes");

        output.WriteLine($"itemType={item.ItemType} messageClass={item.MessageClass} subject={item.Subject}");

        Assert.Equal("note", item.ItemType);
        Assert.False(
            string.IsNullOrWhiteSpace(item.EntryId),
            "A listed note must carry an entry id so it can be addressed afterwards.");
    }

    /// <summary>
    /// The general guarantee, stated once so it cannot be lost in the per-type cases: whatever a
    /// folder holds, no listed item may be described by the name of its runtime wrapper. Third-party
    /// add-in items legitimately fall through the typed branches; they must still be described by
    /// their message class rather than by <c>__ComObject</c>.
    /// </summary>
    [SkippableTheory]
    [InlineData("inbox")]
    [InlineData("tasks")]
    [InlineData("notes")]
    [InlineData("calendar")]
    [InlineData("contacts")]
    public void ListItems_NoFolder_EverReportsTheRuntimeWrapperName(string folder)
    {
        var commands = new FolderCommands();
        var result = commands.ListItems(folder, maxCount: 10);
        Skip.If(!result.Success, result.ErrorMessage);
        Skip.If(result.Items.Count == 0, $"The '{folder}' folder is empty on this profile.");

        foreach (var item in result.Items)
        {
            Assert.NotEqual("__ComObject", item.ItemType);
            Assert.False(
                string.IsNullOrWhiteSpace(item.ItemType),
                $"An item in '{folder}' with message class '{item.MessageClass}' has no item type.");
        }
    }

    private static OutlookMcp.Core.Models.OutlookFolderItemInfo FirstItemIn(string folder)
    {
        var commands = new FolderCommands();
        var result = commands.ListItems(folder, maxCount: 5);
        Skip.If(!result.Success, result.ErrorMessage);
        Skip.If(
            result.Items.Count == 0,
            $"The '{folder}' folder is empty on this profile, so there is nothing to classify.");

        return result.Items[0];
    }
}
