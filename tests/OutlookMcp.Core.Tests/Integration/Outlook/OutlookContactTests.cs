using OutlookMcp.Core.Commands.Contact;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Contact items: list, read, create, update, delete, and distribution lists (#14).
///
/// <para>
/// Contacts were the last unmerged slice of <c>feature/outlook-parity-slices</c>. The stranded
/// implementation listed a folder by walking <c>Items</c> and doing
/// <c>rawItem as ContactItem; if (contact == null) continue;</c>, which is the failure mode this
/// project keeps rediscovering: the folder reports <c>totalItemCount: 83</c>, the caller receives 82
/// rows, and nothing anywhere says what happened to the 83rd. On the developer's real Contacts
/// folder that missing item is a distribution list - a group of people, silently absent from a tool
/// whose whole job is answering "who do I know". The same shape of bug was fixed for mail threads in
/// #112, so it is asserted here rather than shipped again.
/// </para>
///
/// <para>
/// <b>Mutation safety.</b> Contacts created here are named with a GUID scratch prefix, live in the
/// default Contacts folder, and are deleted in <c>finally</c> with a prefix sweep afterwards. No
/// pre-existing contact is ever updated or deleted: every destructive assertion below is made
/// against an item the test created moments earlier.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Contact")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookContactTests(ITestOutputHelper output)
{
    /// <summary>
    /// The baseline: contacts come back, and every one carries the id needed to address it. A name
    /// alone is not a handle - two people share a name, and 1 of the 83 items in the developer's
    /// folder has no name at all.
    /// </summary>
    [SkippableFact]
    public void List_ReturnsContactsThatCanEachBeAddressedById()
    {
        EnsureOutlookAvailable();

        var result = new ContactCommands().List();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Contacts);

        foreach (var contact in result.Contacts)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(contact.EntryId),
                $"Contact '{contact.DisplayName}' arrived without an entry id.");
        }

        output.WriteLine(
            $"{result.ReturnedCount} of {result.TotalItemCount} item(s) from {result.FolderPath}");
    }

    /// <summary>
    /// The reason this file exists. A Contacts folder holds distribution lists as well as people,
    /// and a listing that quietly discards them is wrong in the one way a caller cannot detect.
    ///
    /// <para>
    /// The assertion is deliberately an accounting identity rather than "a distribution list is
    /// present", because a profile with no distribution lists is legitimate and a test that only
    /// passes on the developer's mailbox is a test that stops testing. What must hold either way is
    /// that everything scanned is accounted for: returned contacts, plus returned distribution
    /// lists, plus items explicitly reported as skipped. Nothing may vanish.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void List_AccountsForEveryItemItScanned_IncludingDistributionLists()
    {
        EnsureOutlookAvailable();

        // A limit above the folder size, so the scan really does reach every item and the identity
        // below is about correctness rather than about paging.
        var result = new ContactCommands().List(maxCount: 1000);

        Assert.True(result.Success, result.ErrorMessage);

        int accountedFor = result.Contacts.Count + result.DistributionLists.Count + result.SkippedItemCount;

        Assert.Equal(result.ScannedItemCount, accountedFor);

        foreach (var list in result.DistributionLists)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(list.EntryId),
                $"Distribution list '{list.DisplayName}' arrived without an entry id.");
        }

        output.WriteLine(
            $"total={result.TotalItemCount} scanned={result.ScannedItemCount} "
            + $"contacts={result.Contacts.Count} distributionLists={result.DistributionLists.Count} "
            + $"skipped={result.SkippedItemCount} truncated={result.Truncated}");
    }

    /// <summary>
    /// A row with an empty label is a row a caller cannot act on or show. The developer's folder
    /// contains a contact whose <c>FullName</c> is blank, so falling back to company or email
    /// address is the difference between a usable listing and one with holes in it.
    /// </summary>
    [SkippableFact]
    public void List_NeverReturnsAContactWithNothingToDisplay()
    {
        EnsureOutlookAvailable();

        var result = new ContactCommands().List(maxCount: 1000);

        Assert.True(result.Success, result.ErrorMessage);

        foreach (var contact in result.Contacts)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(contact.DisplayName),
                $"A contact with entry id '{contact.EntryId}' has no display name, "
                + "no company and no email address to fall back to.");
        }
    }

    /// <summary>
    /// The listing and the reader have to agree. An id handed out by <c>list</c> that <c>read</c>
    /// cannot resolve would make the two halves of the tool useless together.
    /// </summary>
    [SkippableFact]
    public void Read_ResolvesAnIdHandedOutByList()
    {
        EnsureOutlookAvailable();

        var commands = new ContactCommands();

        var listing = commands.List(maxCount: 5, includeBodyPreview: false);
        Assert.True(listing.Success, listing.ErrorMessage);
        Skip.If(listing.Contacts.Count == 0, "This profile has no contacts to read.");

        var expected = listing.Contacts[0];

        var read = commands.Read(expected.EntryId, expected.StoreId, useActiveContact: false);

        Assert.True(read.Success, read.ErrorMessage);
        Assert.True(read.HasItem);
        Assert.Equal(expected.EntryId, read.EntryId);
        Assert.Equal(expected.DisplayName, read.DisplayName);

        output.WriteLine($"Read back: {read.DisplayName} | {read.FolderPath}");
    }

    /// <summary>
    /// The full mutation lifecycle, verified from the outside at every step.
    ///
    /// <para>
    /// Each stage is confirmed by re-reading the item rather than by trusting the return value of
    /// the call that made the change. A <c>create</c> that reported success without saving anything,
    /// or an <c>update</c> that reported success while writing to a detached copy, would pass any
    /// assertion made only on its own result - and that is precisely the class of bug this project
    /// keeps finding.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void Create_Update_Delete_ReallyChangeTheContactsFolder()
    {
        EnsureOutlookAvailable();

        var commands = new ContactCommands();
        string lastName = ScratchName();

        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = commands.Create(
                lastName: lastName,
                firstName: "Scratch",
                companyName: "OutlookMcp integration test",
                email1Address: "scratch@example.invalid");

            Assert.True(created.Success, created.ErrorMessage);
            Assert.False(string.IsNullOrWhiteSpace(created.EntryId), "create returned no entry id.");

            entryId = created.EntryId;
            storeId = created.StoreId;
            output.WriteLine($"Created: {entryId}");

            var afterCreate = commands.Read(entryId, storeId, useActiveContact: false);
            Assert.True(afterCreate.Success, afterCreate.ErrorMessage);
            Assert.Equal(lastName, afterCreate.LastName);
            Assert.Equal("scratch@example.invalid", afterCreate.Email1Address);

            var updated = commands.Update(
                entryId,
                storeId,
                jobTitle: "Updated by test",
                useActiveContact: false);

            Assert.True(updated.Success, updated.ErrorMessage);

            var afterUpdate = commands.Read(entryId, storeId, useActiveContact: false);
            Assert.True(afterUpdate.Success, afterUpdate.ErrorMessage);
            Assert.Equal("Updated by test", afterUpdate.JobTitle);

            // The field that was not passed to update must survive it. An update implemented as
            // "write every parameter" would blank this, and a test that only checked the changed
            // field would not notice.
            Assert.Equal("scratch@example.invalid", afterUpdate.Email1Address);
        }
        finally
        {
            if (entryId != null)
            {
                var deleted = commands.Delete(entryId, storeId);
                output.WriteLine($"Delete: success={deleted.Success} {deleted.ErrorMessage}");
                Assert.True(deleted.Success, deleted.ErrorMessage);
            }

            SweepScratchContacts(commands);
        }

        if (entryId != null)
        {
            var afterDelete = commands.Read(entryId, storeId, useActiveContact: false);
            Assert.False(
                afterDelete.Success && afterDelete.HasItem,
                "The contact is still readable after delete reported success.");
        }
    }

    /// <summary>
    /// Update has to refuse an id it cannot resolve rather than inventing a contact. Silently
    /// creating one would be the worst possible outcome of a typo.
    /// </summary>
    [SkippableFact]
    public void Update_RefusesAnIdThatDoesNotResolve()
    {
        EnsureOutlookAvailable();

        var commands = new ContactCommands();
        int before = commands.List(maxCount: 1000).TotalItemCount;

        var result = commands.Update(
            entryId: "0000000000000000000000000000000000000000000000",
            jobTitle: "should never be written",
            useActiveContact: false);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));

        int after = commands.List(maxCount: 1000).TotalItemCount;
        Assert.Equal(before, after);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// The distinction the workflow guidance promises to callers: an omitted field is left alone,
    /// an empty string clears the stored value. Both halves are asserted on the same contact,
    /// because an implementation that treated "" as "not passed" would satisfy the first half and
    /// silently break the second - and that is the half a caller reaches for when asked to remove
    /// someone's old phone number.
    /// </summary>
    [SkippableFact]
    public void Update_ClearsAFieldPassedAsEmpty_AndLeavesOmittedFieldsAlone()
    {
        EnsureOutlookAvailable();

        var commands = new ContactCommands();
        string lastName = ScratchName();

        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = commands.Create(
                lastName: lastName,
                firstName: "Scratch",
                jobTitle: "Job title that must survive",
                businessTelephoneNumber: "+1 555 0100",
                email1Address: "scratch@example.invalid");

            Assert.True(created.Success, created.ErrorMessage);
            entryId = created.EntryId;
            storeId = created.StoreId;

            var beforeClear = commands.Read(entryId, storeId, useActiveContact: false);
            Assert.True(beforeClear.Success, beforeClear.ErrorMessage);

            // Not an equality check: Outlook canonicalises phone numbers on write, so "+1 555 0100"
            // comes back as "+1 (555) 0100". All this step needs to establish is that there is
            // something there for the clear to remove.
            Assert.False(
                string.IsNullOrEmpty(beforeClear.BusinessTelephoneNumber),
                "The phone number was not stored, so clearing it would prove nothing.");
            output.WriteLine($"Stored phone number reads back as '{beforeClear.BusinessTelephoneNumber}'.");

            var cleared = commands.Update(
                entryId,
                storeId,
                businessTelephoneNumber: string.Empty,
                useActiveContact: false);

            Assert.True(cleared.Success, cleared.ErrorMessage);

            var afterClear = commands.Read(entryId, storeId, useActiveContact: false);
            Assert.True(afterClear.Success, afterClear.ErrorMessage);

            Assert.True(
                string.IsNullOrEmpty(afterClear.BusinessTelephoneNumber),
                $"An empty string did not clear the phone number; it is still '{afterClear.BusinessTelephoneNumber}'.");

            // The other fields were not passed, so they must be untouched by the clear.
            Assert.Equal("Job title that must survive", afterClear.JobTitle);
            Assert.Equal("scratch@example.invalid", afterClear.Email1Address);
        }
        finally
        {
            if (entryId != null)
            {
                var deleted = commands.Delete(entryId, storeId);
                Assert.True(deleted.Success, deleted.ErrorMessage);
            }

            SweepScratchContacts(commands);
        }
    }

    /// <summary>
    /// Removes anything this file created that a <c>finally</c> block failed to remove. Reported
    /// loudly rather than swallowed: leftovers here are real rows in a real address book.
    /// </summary>
    private void SweepScratchContacts(ContactCommands commands)
    {
        var listing = commands.List(maxCount: 1000);
        if (!listing.Success)
        {
            output.WriteLine($"Sweep could not list contacts: {listing.ErrorMessage}");
            return;
        }

        var leftovers = listing.Contacts
            .Where(c => c.LastName?.StartsWith(ScratchPrefix, StringComparison.Ordinal) == true)
            .ToList();

        foreach (var contact in leftovers)
        {
            var deleted = commands.Delete(contact.EntryId, contact.StoreId);
            output.WriteLine($"Sweep: {contact.DisplayName} -> success={deleted.Success} {deleted.ErrorMessage}");
        }

        if (leftovers.Count > 0)
        {
            output.WriteLine($"SWEEP removed {leftovers.Count} leftover scratch contact(s).");
        }
    }

    private const string ScratchPrefix = "mcp-test-";

    private static string ScratchName() => $"{ScratchPrefix}{Guid.NewGuid():N}";

    private static void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
