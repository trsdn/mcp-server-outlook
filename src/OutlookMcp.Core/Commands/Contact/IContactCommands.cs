using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Contact;

[ServiceCategory("contact")]
[McpTool("contact", Title = "Outlook Contact Operations", Destructive = false, Category = "contact",
    Description = "Inspect and change Outlook contacts without opening a persistent session. "
    + "Use list to enumerate the default Contacts folder or an explicit Outlook folder path. "
    + "A Contacts folder holds distribution lists as well as people: those are returned separately in "
    + "distributionLists rather than dropped, and contacts, distribution lists and skippedItemCount "
    + "together always account for every item scanned. "
    + "Use read to inspect one contact by entry id, or the contact currently open or selected in Outlook. "
    + "Use create to save a new contact, update to change named fields on an existing one - fields that are "
    + "not passed are left alone - and delete to remove one. delete needs no confirmation in the ordinary case, "
    + "because Outlook moves the contact to Deleted Items where the user can restore it; deleting a contact that "
    + "is already in Deleted Items is permanent and requires confirm=true. "
    + "Every contact carries an entryId; names are not unique and some contacts have no name at all, "
    + "so entryId is the only reliable handle.")]
public interface IContactCommands
{
    [ServiceAction("list")]
    ContactListResult List(
        string? folder = null,
        int maxCount = 25,
        bool includeBodyPreview = false);

    [ServiceAction("read")]
    ContactItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveContact = true);

    [ServiceAction("create", Destructive = true)]
    ContactMutationResult Create(
        string? folder = null,
        string? firstName = null,
        string? lastName = null,
        string? companyName = null,
        string? jobTitle = null,
        string? email1Address = null,
        string? email2Address = null,
        string? businessTelephoneNumber = null,
        string? mobileTelephoneNumber = null,
        string? body = null,
        bool display = false);

    [ServiceAction("update", Destructive = true)]
    ContactMutationResult Update(
        string? entryId = null,
        string? storeId = null,
        string? firstName = null,
        string? lastName = null,
        string? companyName = null,
        string? jobTitle = null,
        string? email1Address = null,
        string? email2Address = null,
        string? businessTelephoneNumber = null,
        string? mobileTelephoneNumber = null,
        string? body = null,
        bool useActiveContact = true);

    /// <summary>
    /// Deletes a contact.
    ///
    /// <para>
    /// An ordinary delete moves the contact to Deleted Items, where the user can restore it, so it
    /// is not gated. Deleting a contact that is already in Deleted Items destroys it and requires
    /// <paramref name="confirm"/>.
    /// </para>
    /// </summary>
    /// <param name="confirm">Required only when the contact is already in Deleted Items. An ordinary delete ignores it.</param>
    [ServiceAction("delete", Destructive = true)]
    ContactMutationResult Delete(
        string? entryId = null,
        string? storeId = null,
        bool useActiveContact = false,
        bool confirm = false);
}
