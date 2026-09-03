using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Contact;

/// <summary>
/// Outlook contact operations (#14).
///
/// <para>
/// A Contacts folder is not a folder of contacts. It holds <c>ContactItem</c> and
/// <c>DistListItem</c>, and can hold anything else a user has filed there. Everything in
/// <see cref="List"/> is built around that: items are classified rather than filtered, and the
/// counts returned add up, so a caller can tell the difference between "this folder holds 82
/// people" and "this folder holds 83 things, 82 of which are people".
/// </para>
/// </summary>
public class ContactCommands : IContactCommands
{
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> FolderAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["contacts"] = Outlook.OlDefaultFolders.olFolderContacts,
            ["contact"] = Outlook.OlDefaultFolders.olFolderContacts,
            ["current"] = Outlook.OlDefaultFolders.olFolderContacts
        };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ContactListResult List(
        string? folder = null,
        int maxCount = 25,
        bool includeBodyPreview = false)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);

        return OutlookInteropRunner.Execute(
            "OutlookContactList",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? contactFolder = null;
                Outlook.Items? items = null;

                try
                {
                    contactFolder = ResolveContactFolder(application, session, folder, ref explorer);
                    if (contactFolder == null)
                    {
                        return new ContactListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = contactFolder.Items;
                    int totalItemCount = SafeGetInt(() => items.Count);
                    TrySortByLastModification(items);

                    var result = new ContactListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => contactFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(contactFolder),
                        TotalItemCount = totalItemCount
                    };

                    int scanned = 0;

                    for (int index = 1; index <= totalItemCount; index++)
                    {
                        if (result.Contacts.Count + result.DistributionLists.Count >= boundedMaxCount)
                        {
                            break;
                        }

                        object? rawItem = null;

                        try
                        {
                            rawItem = items[index];
                            scanned++;

                            switch (rawItem)
                            {
                                case Outlook.ContactItem contact:
                                    result.Contacts.Add(CreateContactSummary(contact, includeBodyPreview));
                                    break;

                                case Outlook.DistListItem distributionList:
                                    result.DistributionLists.Add(CreateDistributionListInfo(distributionList));
                                    break;

                                default:
                                    // Something filed into a Contacts folder that is neither a person
                                    // nor a group. Counted rather than ignored so the totals add up.
                                    result.SkippedItemCount++;
                                    break;
                            }
                        }
                        catch (COMException)
                        {
                            // The item exists but could not be read - a corrupt row, or one the
                            // Object Model Guard refuses. Counted, because the alternative is a
                            // listing that is quietly short.
                            scanned++;
                            result.SkippedItemCount++;
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ScannedItemCount = scanned;
                    result.ReturnedCount = result.Contacts.Count;
                    result.Truncated = scanned < totalItemCount;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref contactFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new ContactListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook contacts: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ContactItemResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveContact = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookContactRead",
            (application, session) =>
            {
                var resolved = ResolveContactItem(application, session, entryId, storeId, useActiveContact);

                try
                {
                    if (resolved.Contact == null)
                    {
                        return new ContactItemResult
                        {
                            Success = true,
                            HasItem = false
                        };
                    }

                    return CreateContactItemResult(resolved.Contact);
                }
                finally
                {
                    resolved.Release();
                }
            },
            ex => new ContactItemResult
            {
                Success = false,
                HasItem = false,
                ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                    ? $"Failed to inspect the active Outlook contact: {ex.Message}"
                    : $"Failed to inspect the requested Outlook contact: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ContactMutationResult Create(
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
        bool display = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookContactCreate",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? contactFolder = null;
                Outlook.Items? items = null;
                object? createdItem = null;
                Outlook.ContactItem? contact = null;

                try
                {
                    contactFolder = ResolveContactFolder(application, session, folder, ref explorer);
                    if (contactFolder == null)
                    {
                        return new ContactMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = contactFolder.Items;
                    createdItem = items.Add(Outlook.OlItemType.olContactItem);
                    contact = createdItem as Outlook.ContactItem;

                    if (contact == null)
                    {
                        return new ContactMutationResult
                        {
                            Success = false,
                            ErrorMessage = "Outlook did not return a contact item for the new contact."
                        };
                    }

                    ApplyContactUpdates(
                        contact,
                        firstName,
                        lastName,
                        companyName,
                        jobTitle,
                        email1Address,
                        email2Address,
                        businessTelephoneNumber,
                        mobileTelephoneNumber,
                        body);

                    contact.Save();

                    if (display)
                    {
                        contact.Display(false);
                    }

                    var result = CreateContactMutationResult(contact, "Saved Outlook contact.");
                    result.Saved = true;
                    result.Displayed = display;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref contact);
                    OutlookInteropRunner.ReleaseComObject(ref createdItem);
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref contactFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new ContactMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create the Outlook contact: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ContactMutationResult Update(
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
        bool useActiveContact = true)
    {
        bool hasUpdates =
            firstName != null ||
            lastName != null ||
            companyName != null ||
            jobTitle != null ||
            email1Address != null ||
            email2Address != null ||
            businessTelephoneNumber != null ||
            mobileTelephoneNumber != null ||
            body != null;

        if (!hasUpdates)
        {
            return new ContactMutationResult
            {
                Success = false,
                ErrorMessage = "At least one contact field must be provided for update."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookContactUpdate",
            (application, session) =>
            {
                var resolved = ResolveContactItem(application, session, entryId, storeId, useActiveContact);

                try
                {
                    if (resolved.Contact == null)
                    {
                        return new ContactMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedMessage(entryId, "update")
                        };
                    }

                    ApplyContactUpdates(
                        resolved.Contact,
                        firstName,
                        lastName,
                        companyName,
                        jobTitle,
                        email1Address,
                        email2Address,
                        businessTelephoneNumber,
                        mobileTelephoneNumber,
                        body);

                    resolved.Contact.Save();

                    var result = CreateContactMutationResult(resolved.Contact, "Updated Outlook contact.");
                    result.Saved = true;
                    result.Updated = true;
                    return result;
                }
                finally
                {
                    resolved.Release();
                }
            },
            ex => new ContactMutationResult
            {
                Success = false,
                ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                    ? $"Failed to update the active Outlook contact: {ex.Message}"
                    : $"Failed to update the requested Outlook contact: {ex.Message}"
            });
    }

    /// <summary>
    /// Deletes a contact.
    ///
    /// <para>
    /// <c>useActiveContact</c> defaults to false here, unlike read and update. Falling back to
    /// whatever happens to be selected in Outlook is a convenience when reading and a hazard when
    /// deleting: a delete call with a mistyped id would otherwise remove a different contact
    /// entirely and report success.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ContactMutationResult Delete(
        string? entryId = null,
        string? storeId = null,
        bool useActiveContact = false)
    {
        if (string.IsNullOrWhiteSpace(entryId) && !useActiveContact)
        {
            return new ContactMutationResult
            {
                Success = false,
                ErrorMessage = "An entryId is required to delete a contact. "
                    + "Pass useActiveContact: true to delete the contact currently open or selected in Outlook."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookContactDelete",
            (application, session) =>
            {
                var resolved = ResolveContactItem(application, session, entryId, storeId, useActiveContact);

                try
                {
                    if (resolved.Contact == null)
                    {
                        return new ContactMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedMessage(entryId, "delete")
                        };
                    }

                    // Read the identifying fields before the delete, because afterwards the item is
                    // gone and every property on it throws.
                    var result = CreateContactMutationResult(resolved.Contact, "Deleted Outlook contact.");

                    resolved.Contact.Delete();
                    result.Deleted = true;
                    return result;
                }
                finally
                {
                    resolved.Release();
                }
            },
            ex => new ContactMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to delete the Outlook contact: {ex.Message}"
            });
    }

    /// <summary>
    /// A resolved contact together with the COM objects that had to be held to reach it. Returning
    /// them as one value keeps the release list in one place: the caller cannot forget the explorer
    /// it never asked for.
    /// </summary>
    private sealed class ResolvedContact
    {
        public Outlook.ContactItem? Contact { get; set; }

        public Outlook.Inspector? Inspector { get; set; }

        public Outlook.Explorer? Explorer { get; set; }

        public Outlook.Selection? Selection { get; set; }

        public object? RawItem { get; set; }

        public void Release()
        {
            Outlook.ContactItem? contact = Contact;
            Outlook.Inspector? inspector = Inspector;
            Outlook.Explorer? explorer = Explorer;
            Outlook.Selection? selection = Selection;
            object? rawItem = RawItem;

            OutlookInteropRunner.ReleaseComObject(ref contact);
            OutlookInteropRunner.ReleaseComObject(ref rawItem);
            OutlookInteropRunner.ReleaseComObject(ref selection);
            OutlookInteropRunner.ReleaseComObject(ref explorer);
            OutlookInteropRunner.ReleaseComObject(ref inspector);

            Contact = null;
            Inspector = null;
            Explorer = null;
            Selection = null;
            RawItem = null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ResolvedContact ResolveContactItem(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? entryId,
        string? storeId,
        bool useActiveContact)
    {
        var resolved = new ResolvedContact();

        if (!string.IsNullOrWhiteSpace(entryId))
        {
            try
            {
                resolved.RawItem = session.GetItemFromID(
                    entryId,
                    string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
                resolved.Contact = resolved.RawItem as Outlook.ContactItem;
            }
            catch (COMException)
            {
                // An id that does not resolve. Reported by the caller as a refusal rather than
                // silently falling through to whatever is selected in the UI.
            }

            return resolved;
        }

        if (!useActiveContact)
        {
            return resolved;
        }

        resolved.Inspector = application.ActiveInspector();
        if (resolved.Inspector != null)
        {
            resolved.RawItem = resolved.Inspector.CurrentItem;
            if (resolved.RawItem is Outlook.ContactItem openContact)
            {
                resolved.Contact = openContact;
                return resolved;
            }

            object? notAContact = resolved.RawItem;
            OutlookInteropRunner.ReleaseComObject(ref notAContact);
            resolved.RawItem = null;
        }

        resolved.Explorer = application.ActiveExplorer();
        if (resolved.Explorer != null)
        {
            resolved.Selection = resolved.Explorer.Selection;
            if (resolved.Selection != null && resolved.Selection.Count > 0)
            {
                resolved.RawItem = resolved.Selection[1];
                if (resolved.RawItem is Outlook.ContactItem selectedContact)
                {
                    resolved.Contact = selectedContact;
                }
            }
        }

        return resolved;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ContactItemResult CreateContactItemResult(Outlook.ContactItem contact)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = contact.Parent as Outlook.MAPIFolder;

            string? fullName = SafeGet(() => contact.FullName);
            string? companyName = SafeGet(() => contact.CompanyName);
            string? email1Address = SafeGet(() => contact.Email1Address);

            return new ContactItemResult
            {
                Success = true,
                HasItem = true,
                EntryId = SafeGet(() => contact.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                DisplayName = BuildDisplayName(fullName, companyName, email1Address),
                FullName = NullIfBlank(fullName),
                FirstName = NullIfBlank(SafeGet(() => contact.FirstName)),
                LastName = NullIfBlank(SafeGet(() => contact.LastName)),
                CompanyName = NullIfBlank(companyName),
                JobTitle = NullIfBlank(SafeGet(() => contact.JobTitle)),
                Email1Address = NullIfBlank(email1Address),
                Email2Address = NullIfBlank(SafeGet(() => contact.Email2Address)),
                BusinessTelephoneNumber = NullIfBlank(SafeGet(() => contact.BusinessTelephoneNumber)),
                MobileTelephoneNumber = NullIfBlank(SafeGet(() => contact.MobileTelephoneNumber)),
                FolderPath = OutlookInteropRunner.GetFolderPath(parentFolder),
                BodyPreview = NullIfBlank(OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => contact.Body)))
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ContactSummaryInfo CreateContactSummary(Outlook.ContactItem contact, bool includeBodyPreview)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = contact.Parent as Outlook.MAPIFolder;

            string? fullName = SafeGet(() => contact.FullName);
            string? companyName = SafeGet(() => contact.CompanyName);
            string? email1Address = SafeGet(() => contact.Email1Address);

            return new ContactSummaryInfo
            {
                EntryId = SafeGet(() => contact.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                DisplayName = BuildDisplayName(fullName, companyName, email1Address),
                FullName = NullIfBlank(fullName),
                LastName = NullIfBlank(SafeGet(() => contact.LastName)),
                CompanyName = NullIfBlank(companyName),
                Email1Address = NullIfBlank(email1Address),
                BusinessTelephoneNumber = NullIfBlank(SafeGet(() => contact.BusinessTelephoneNumber)),
                MobileTelephoneNumber = NullIfBlank(SafeGet(() => contact.MobileTelephoneNumber)),
                BodyPreview = includeBodyPreview
                    ? NullIfBlank(OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => contact.Body)))
                    : null
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ContactDistributionListInfo CreateDistributionListInfo(Outlook.DistListItem distributionList)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = distributionList.Parent as Outlook.MAPIFolder;

            string? name = SafeGet(() => distributionList.DLName) ?? SafeGet(() => distributionList.Subject);

            return new ContactDistributionListInfo
            {
                EntryId = SafeGet(() => distributionList.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                DisplayName = string.IsNullOrWhiteSpace(name) ? "(unnamed distribution list)" : name,
                MemberCount = SafeGetInt(() => distributionList.MemberCount)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ContactMutationResult CreateContactMutationResult(Outlook.ContactItem contact, string message)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = contact.Parent as Outlook.MAPIFolder;

            string? fullName = SafeGet(() => contact.FullName);
            string? companyName = SafeGet(() => contact.CompanyName);
            string? email1Address = SafeGet(() => contact.Email1Address);

            return new ContactMutationResult
            {
                Success = true,
                Message = message,
                EntryId = SafeGet(() => contact.EntryID),
                StoreId = SafeGet(() => parentFolder?.StoreID),
                DisplayName = BuildDisplayName(fullName, companyName, email1Address),
                FullName = NullIfBlank(fullName),
                CompanyName = NullIfBlank(companyName),
                JobTitle = NullIfBlank(SafeGet(() => contact.JobTitle)),
                Email1Address = NullIfBlank(email1Address),
                Email2Address = NullIfBlank(SafeGet(() => contact.Email2Address)),
                BusinessTelephoneNumber = NullIfBlank(SafeGet(() => contact.BusinessTelephoneNumber)),
                MobileTelephoneNumber = NullIfBlank(SafeGet(() => contact.MobileTelephoneNumber)),
                FolderPath = OutlookInteropRunner.GetFolderPath(parentFolder)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    /// <summary>
    /// Produces a label that is never blank.
    ///
    /// <para>
    /// Contacts really do exist with no name: an address harvested from a mail header, or a company
    /// record. Returning an empty string for those makes a listing that a caller cannot render or
    /// disambiguate, so company and email address are used in turn before admitting there is
    /// nothing.
    /// </para>
    /// </summary>
    private static string BuildDisplayName(string? fullName, string? companyName, string? email1Address)
    {
        if (!string.IsNullOrWhiteSpace(fullName))
        {
            return fullName;
        }

        if (!string.IsNullOrWhiteSpace(companyName))
        {
            return companyName;
        }

        if (!string.IsNullOrWhiteSpace(email1Address))
        {
            return email1Address;
        }

        return "(contact with no name)";
    }

    private static void ApplyContactUpdates(
        Outlook.ContactItem contact,
        string? firstName,
        string? lastName,
        string? companyName,
        string? jobTitle,
        string? email1Address,
        string? email2Address,
        string? businessTelephoneNumber,
        string? mobileTelephoneNumber,
        string? body)
    {
        // Only fields that were actually passed are written. A parameter left unset must leave the
        // stored value alone, otherwise changing a job title would blank a phone number.
        if (firstName != null)
        {
            contact.FirstName = firstName;
        }

        if (lastName != null)
        {
            contact.LastName = lastName;
        }

        if (companyName != null)
        {
            contact.CompanyName = companyName;
        }

        if (jobTitle != null)
        {
            contact.JobTitle = jobTitle;
        }

        if (email1Address != null)
        {
            contact.Email1Address = email1Address;
        }

        if (email2Address != null)
        {
            contact.Email2Address = email2Address;
        }

        if (businessTelephoneNumber != null)
        {
            contact.BusinessTelephoneNumber = businessTelephoneNumber;
        }

        if (mobileTelephoneNumber != null)
        {
            contact.MobileTelephoneNumber = mobileTelephoneNumber;
        }

        if (body != null)
        {
            contact.Body = body;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? ResolveContactFolder(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? folder,
        ref Outlook.Explorer? explorer)
        => OutlookInteropRunner.ResolveFolder(
            application,
            session,
            string.IsNullOrWhiteSpace(folder) ? "contacts" : folder,
            FolderAliases,
            ref explorer);

    private static string BuildUnknownFolderMessage(string? folder)
    {
        const string supportedFolders = "current, contacts, or an Outlook folder path";
        return string.IsNullOrWhiteSpace(folder)
            ? $"Could not resolve the Outlook contact folder. Supported folder values: {supportedFolders}."
            : $"Unsupported Outlook contact folder '{folder}'. Supported folder values: {supportedFolders}.";
    }

    private static string BuildUnresolvedMessage(string? entryId, string operation)
        => string.IsNullOrWhiteSpace(entryId)
            ? $"Could not resolve an active Outlook contact to {operation}. "
              + "Open or select a contact in Outlook, or pass an entryId from contact list."
            : $"Could not resolve an Outlook contact with entryId '{entryId}' to {operation}. "
              + "The id may belong to a deleted item, to a different store, or to an item that is not a contact.";

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void TrySortByLastModification(Outlook.Items items)
    {
        try
        {
            // Contacts have no received time, so the mail ordering used elsewhere does not apply.
            items.Sort("[LastModificationTime]", true);
        }
        catch (COMException)
        {
            // Some folders refuse to sort. Store order is then the honest answer.
        }
    }

    private static string? NullIfBlank(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value;

    private static string? SafeGet(Func<string?> getter)
    {
        try
        {
            return getter();
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static int SafeGetInt(Func<int> getter)
    {
        try
        {
            return getter();
        }
        catch (COMException)
        {
            return 0;
        }
    }
}
