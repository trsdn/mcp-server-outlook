using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.AddressBook;

/// <summary>
/// Address book lookup and recipient resolution (#15).
///
/// <para>
/// The single fact this class exists to get right: for an Exchange entry,
/// <c>AddressEntry.Address</c> is an X500 legacyExchangeDN
/// (<c>/o=ExchangeLabs/ou=.../cn=Recipients/cn=...</c>), not an email address. It is a string, it
/// serialises cleanly, and mail sent to it goes nowhere. The real address has to be fetched
/// through <c>GetExchangeUser().PrimarySmtpAddress</c> - or
/// <c>GetExchangeDistributionList().PrimarySmtpAddress</c>, which is a different call on a
/// different COM type for a group - with <c>PR_SMTP_ADDRESS</c> as a fallback when the directory
/// is offline or the entry is neither. Both values are reported, and which route produced the
/// address is reported with them.
/// </para>
///
/// <para>
/// Everything here is read-only, and nearly all of it is Object Model Guard territory: every
/// property and method on <c>Recipient</c> is protected, along with <c>AddressEntry.Address</c>,
/// <c>GetExchangeUser</c>, <c>ExchangeUser.PrimarySmtpAddress</c> and every cursor method on
/// <c>AddressEntries</c>. A denial cannot be answered programmatically, so it is surfaced -
/// whole-call denials through the runner's classification, per-property denials through
/// <c>accessDenied</c>.
/// </para>
/// </summary>
public class AddressBookCommands : IAddressBookCommands
{
    /// <summary>
    /// <c>PR_SMTP_ADDRESS</c>, Unicode. Tried before the ANSI form; Outlook coerces between
    /// <c>001E</c> and <c>001F</c> for string properties, so either can answer, but a store that
    /// holds only one of them will refuse the other with <c>MAPI_E_NOT_FOUND</c>.
    /// </summary>
    private const string PrSmtpAddressUnicode = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

    /// <summary><c>PR_SMTP_ADDRESS</c>, ANSI. The form Microsoft's own sample uses.</summary>
    private const string PrSmtpAddressAnsi = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

    private const int MaxReturnedEntries = 100;
    private const int MaxScannedEntries = 5000;

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AddressListCollectionResult ListAddressLists(bool includeEntryCount = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookAddressBookListAddressLists",
            (application, session) =>
            {
                Outlook.AddressLists? lists = null;

                try
                {
                    lists = session.AddressLists;
                    int count = SafeGetInt(() => lists.Count);

                    var result = new AddressListCollectionResult { Success = true };

                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.AddressList? list = null;

                        try
                        {
                            list = lists[index];
                            result.AddressLists.Add(DescribeAddressList(list, index, includeEntryCount));
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref list);
                        }
                    }

                    result.Count = result.AddressLists.Count;
                    result.HasGlobalAddressList = result.AddressLists.Exists(
                        l => l.AddressListType == "exchange-global-address-list");
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref lists);
                }
            },
            ex => new AddressListCollectionResult
            {
                Success = false,
                ErrorMessage = $"Failed to list Outlook address books: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AddressResolveResult Resolve(string recipients, bool includeDetails = true)
    {
        List<string> queries = SplitRecipients(recipients);

        if (queries.Count == 0)
        {
            return new AddressResolveResult
            {
                Success = false,
                ErrorMessage = "recipients is required for addressbook.resolve: pass one or more "
                    + "display names, Exchange aliases or email addresses, separated by semicolons."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookAddressBookResolve",
            (application, session) =>
            {
                var result = new AddressResolveResult
                {
                    Success = true,
                    RequestedCount = queries.Count
                };

                foreach (string query in queries)
                {
                    result.Recipients.Add(ResolveOne(session, query, includeDetails));
                }

                result.ResolvedCount = result.Recipients.Count(r => r.Resolved);
                result.AllResolved = result.ResolvedCount == result.RequestedCount;
                result.UnresolvedNames = result.Recipients
                    .Where(r => !r.Resolved)
                    .Select(r => r.Query)
                    .ToList();

                return result;
            },
            ex => new AddressResolveResult
            {
                Success = false,
                ErrorMessage = $"Failed to resolve addressees against the Outlook address book: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AddressEntryListResult ListEntries(
        string? addressList = null,
        string? startsWith = null,
        int maxCount = 25,
        int scanLimit = 500,
        bool includeSmtpAddress = true)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, MaxReturnedEntries);
        int boundedScanLimit = Math.Clamp(scanLimit, 1, MaxScannedEntries);
        string? prefix = NullIfBlank(startsWith);

        return OutlookInteropRunner.Execute(
            "OutlookAddressBookListEntries",
            (application, session) =>
            {
                Outlook.AddressList? list = null;
                Outlook.AddressEntries? entries = null;

                try
                {
                    list = ResolveAddressList(session, addressList);

                    if (list == null)
                    {
                        return new AddressEntryListResult
                        {
                            Success = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(addressList)
                                ? "This Outlook profile exposes no address books."
                                : $"No Outlook address book called '{addressList}'. Call "
                                  + "addressbook.list-address-lists to see which books exist."
                        };
                    }

                    var result = new AddressEntryListResult
                    {
                        Success = true,
                        AddressListName = SafeGet(() => list.Name),
                        AddressListType = DescribeAddressListType(list)
                    };

                    entries = list.AddressEntries;
                    int totalCount = SafeGetInt(() => entries.Count);

                    // GetFirst/GetNext is the cursor API Outlook documents for large collections,
                    // and it must be driven from one held reference to the same AddressEntries
                    // object: re-reading list.AddressEntries inside the loop resets the cursor.
                    Outlook.AddressEntry? entry = entries.GetFirst();
                    int scanned = 0;

                    try
                    {
                        while (entry != null)
                        {
                            scanned++;

                            string? name = SafeGet(() => entry.Name);

                            if (!string.IsNullOrWhiteSpace(name)
                                && (prefix == null || name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                            {
                                result.Entries.Add(DescribeAddressEntry(entry, name, includeSmtpAddress));
                            }

                            if (result.Entries.Count >= boundedMaxCount || scanned >= boundedScanLimit)
                            {
                                break;
                            }

                            OutlookInteropRunner.ReleaseComObject(ref entry);
                            entry = entries.GetNext();
                        }
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref entry);
                    }

                    result.ReturnedCount = result.Entries.Count;
                    result.ScannedCount = scanned;
                    result.ScanLimitReached = scanned >= boundedScanLimit && (totalCount == 0 || scanned < totalCount);
                    result.Truncated = totalCount > 0 && scanned < totalCount;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref entries);
                    OutlookInteropRunner.ReleaseComObject(ref list);
                }
            },
            ex => new AddressEntryListResult
            {
                Success = false,
                ErrorMessage = $"Failed to read the Outlook address book: {ex.Message}"
            });
    }

    // ── Resolution ──────────────────────────────────────────────────────────

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ResolvedRecipientInfo ResolveOne(Outlook.NameSpace session, string query, bool includeDetails)
    {
        Outlook.Recipient? recipient = null;
        Outlook.AddressEntry? entry = null;

        try
        {
            var info = new ResolvedRecipientInfo { Query = query };

            recipient = session.CreateRecipient(query);

            // Resolve() answers false for both "no such addressee" and "the name was ambiguous".
            // Outlook's object model offers no way to tell those apart or to list the candidates,
            // so both are reported as unresolved rather than guessed at.
            info.Resolved = recipient.Resolve();
            info.DisplayName = NullIfBlank(SafeGet(() => recipient.Name, "Recipient.Name", info.AccessDenied));

            if (!info.Resolved)
            {
                return info;
            }

            entry = SafeGetComObject(() => recipient.AddressEntry, "Recipient.AddressEntry", info.AccessDenied);
            PopulateFromAddressEntry(info, entry, recipient, includeDetails);
            return info;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref entry);
            OutlookInteropRunner.ReleaseComObject(ref recipient);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void PopulateFromAddressEntry(
        ResolvedRecipientInfo info,
        Outlook.AddressEntry? entry,
        Outlook.Recipient recipient,
        bool includeDetails)
    {
        if (entry == null)
        {
            info.Note = info.AccessDenied.Contains("Recipient.AddressEntry")
                ? "Outlook resolved the name but its security prompt refused the address entry, so "
                  + "no address could be read. This is not the same as the addressee having none."
                : "Outlook resolved the name but returned no address entry for it, so no "
                  + "address could be read.";
            return;
        }

        Outlook.OlAddressEntryUserType? userType = SafeGetUserType(entry);

        info.EntryType = DescribeEntryType(userType);
        info.AddressType = NullIfBlank(SafeGet(() => entry.Type, "AddressEntry.Type", info.AccessDenied));
        info.RawAddress = NullIfBlank(SafeGet(() => entry.Address, "AddressEntry.Address", info.AccessDenied));
        info.IsDistributionList = userType is Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry
            or Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry;

        info.DisplayName ??= NullIfBlank(SafeGet(() => entry.Name));

        switch (userType)
        {
            case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:
            case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry:
                ReadExchangeUser(info, entry, includeDetails);
                break;

            case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                ReadExchangeDistributionList(info, entry, includeDetails);
                break;

            case Outlook.OlAddressEntryUserType.olOutlookContactAddressEntry:
                ReadContact(info, entry);
                break;

            case Outlook.OlAddressEntryUserType.olSmtpAddressEntry:
                if (LooksLikeSmtpAddress(info.RawAddress))
                {
                    info.SmtpAddress = info.RawAddress;
                    info.SmtpAddressSource = "smtp-entry";
                }

                break;
        }

        if (info.SmtpAddress == null)
        {
            // PR_SMTP_ADDRESS off the Recipient itself. This is the route that still answers when
            // GetExchangeUser returns null - which it does for a public folder, an agent, an LDAP
            // entry, and for any entry at all while the client is offline from the directory.
            string? viaProperty = ReadSmtpProperty(() => recipient.PropertyAccessor, info.AccessDenied);

            if (LooksLikeSmtpAddress(viaProperty))
            {
                info.SmtpAddress = viaProperty;
                info.SmtpAddressSource = "property-accessor";
            }
        }

        if (info.SmtpAddress == null && LooksLikeSmtpAddress(info.RawAddress))
        {
            info.SmtpAddress = info.RawAddress;
            info.SmtpAddressSource = "smtp-entry";
        }

        if (info.SmtpAddress == null)
        {
            info.Note = info.AccessDenied.Count > 0
                ? "No email address could be read; Outlook's security prompt refused the properties "
                  + "named in accessDenied."
                : "Outlook resolved this addressee but exposed no email address for it. The "
                  + "rawAddress value is the provider's own identifier and is not mailable.";
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadExchangeUser(ResolvedRecipientInfo info, Outlook.AddressEntry entry, bool includeDetails)
    {
        Outlook.ExchangeUser? user = null;

        try
        {
            user = SafeGetComObject(() => entry.GetExchangeUser(), "AddressEntry.GetExchangeUser", info.AccessDenied);

            if (user == null)
            {
                // Documented: null when the entry is not an Exchange user, and also whenever the
                // client cannot reach the Exchange server. The PropertyAccessor fallback covers it.
                return;
            }

            string? smtp = NullIfBlank(SafeGet(
                () => user.PrimarySmtpAddress, "ExchangeUser.PrimarySmtpAddress", info.AccessDenied));

            if (LooksLikeSmtpAddress(smtp))
            {
                info.SmtpAddress = smtp;
                info.SmtpAddressSource = "exchange-user";
            }

            if (includeDetails)
            {
                info.Alias = NullIfBlank(SafeGet(() => user.Alias, "ExchangeUser.Alias", info.AccessDenied));
                info.JobTitle = NullIfBlank(SafeGet(() => user.JobTitle));
                info.Department = NullIfBlank(SafeGet(() => user.Department));
                info.OfficeLocation = NullIfBlank(SafeGet(() => user.OfficeLocation));
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref user);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadExchangeDistributionList(
        ResolvedRecipientInfo info,
        Outlook.AddressEntry entry,
        bool includeDetails)
    {
        Outlook.ExchangeDistributionList? list = null;

        try
        {
            // A distribution list is a different COM type reached by a different call.
            // GetExchangeUser() on a group entry returns null, so dispatching on the entry type
            // rather than trying one call and hoping is the only correct shape here.
            list = SafeGetComObject(
                () => entry.GetExchangeDistributionList(),
                "AddressEntry.GetExchangeDistributionList",
                info.AccessDenied);

            if (list == null)
            {
                return;
            }

            string? smtp = NullIfBlank(SafeGet(
                () => list.PrimarySmtpAddress, "ExchangeDistributionList.PrimarySmtpAddress", info.AccessDenied));

            if (LooksLikeSmtpAddress(smtp))
            {
                info.SmtpAddress = smtp;
                info.SmtpAddressSource = "exchange-distribution-list";
            }

            if (includeDetails)
            {
                info.Alias = NullIfBlank(SafeGet(
                    () => list.Alias, "ExchangeDistributionList.Alias", info.AccessDenied));
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref list);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadContact(ResolvedRecipientInfo info, Outlook.AddressEntry entry)
    {
        Outlook.ContactItem? contact = null;

        try
        {
            contact = SafeGetComObject(() => entry.GetContact(), "AddressEntry.GetContact", info.AccessDenied);

            if (contact == null)
            {
                return;
            }

            // A contact's Email1Address is itself an X500 DN when the contact points at an
            // Exchange user, so it is accepted only when it actually looks like an email address.
            string? smtp = NullIfBlank(SafeGet(() => contact.Email1Address));

            if (LooksLikeSmtpAddress(smtp))
            {
                info.SmtpAddress = smtp;
                info.SmtpAddressSource = "contact";
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref contact);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static AddressBookEntryInfo DescribeAddressEntry(
        Outlook.AddressEntry entry,
        string displayName,
        bool includeSmtpAddress)
    {
        Outlook.OlAddressEntryUserType? userType = SafeGetUserType(entry);
        var denied = new List<string>();

        var info = new AddressBookEntryInfo
        {
            DisplayName = displayName,
            EntryType = DescribeEntryType(userType),
            AddressType = NullIfBlank(SafeGet(() => entry.Type, "AddressEntry.Type", denied)),
            RawAddress = NullIfBlank(SafeGet(() => entry.Address, "AddressEntry.Address", denied)),
            IsDistributionList = userType is Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry
                or Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry,
            AccessDenied = denied
        };

        if (!includeSmtpAddress)
        {
            return info;
        }

        var resolved = new ResolvedRecipientInfo { Query = displayName, AccessDenied = denied };

        switch (userType)
        {
            case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:
            case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry:
                ReadExchangeUser(resolved, entry, includeDetails: false);
                break;

            case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                ReadExchangeDistributionList(resolved, entry, includeDetails: false);
                break;

            case Outlook.OlAddressEntryUserType.olOutlookContactAddressEntry:
                ReadContact(resolved, entry);
                break;
        }

        if (resolved.SmtpAddress == null)
        {
            string? viaProperty = ReadSmtpProperty(() => entry.PropertyAccessor, denied);

            if (LooksLikeSmtpAddress(viaProperty))
            {
                resolved.SmtpAddress = viaProperty;
                resolved.SmtpAddressSource = "property-accessor";
            }
        }

        if (resolved.SmtpAddress == null && LooksLikeSmtpAddress(info.RawAddress))
        {
            resolved.SmtpAddress = info.RawAddress;
            resolved.SmtpAddressSource = "smtp-entry";
        }

        info.SmtpAddress = resolved.SmtpAddress;
        info.SmtpAddressSource = resolved.SmtpAddressSource;
        return info;
    }

    /// <summary>
    /// Reads <c>PR_SMTP_ADDRESS</c> through a PropertyAccessor, Unicode form first then ANSI.
    /// The accessor is a COM object in its own right and is released here rather than left to the
    /// garbage collector.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? ReadSmtpProperty(Func<Outlook.PropertyAccessor> accessorFactory, List<string> accessDenied)
    {
        Outlook.PropertyAccessor? accessor = null;

        try
        {
            accessor = SafeGetComObject(accessorFactory, "PropertyAccessor", accessDenied);

            if (accessor == null)
            {
                return null;
            }

            return NullIfBlank(SafeGet(
                       () => accessor.GetProperty(PrSmtpAddressUnicode) as string,
                       "PR_SMTP_ADDRESS",
                       accessDenied))
                ?? NullIfBlank(SafeGet(() => accessor.GetProperty(PrSmtpAddressAnsi) as string));
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref accessor);
        }
    }

    // ── Address list plumbing ───────────────────────────────────────────────

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static AddressListInfo DescribeAddressList(Outlook.AddressList list, int index, bool includeEntryCount)
    {
        var info = new AddressListInfo
        {
            Name = NullIfBlank(SafeGet(() => list.Name)) ?? $"(address list {index})",
            Index = index,
            AddressListType = DescribeAddressListType(list),
            IsReadOnly = SafeGetBool(() => list.IsReadOnly),
            IsInitialAddressList = SafeGetBool(() => list.IsInitialAddressList)
        };

        if (!includeEntryCount)
        {
            return info;
        }

        Outlook.AddressEntries? entries = null;

        try
        {
            entries = list.AddressEntries;
            info.EntryCount = entries.Count;
        }
        catch (COMException ex)
        {
            // An address book that declines to be counted is a fact worth reporting, not a reason
            // to fail the whole listing: an LDAP directory that is not currently reachable is the
            // usual cause, and every other book in the profile is still perfectly describable.
            info.Note = OutlookInteropRunner.IsObjectModelGuardDenial(ex)
                ? "Outlook's security prompt refused the entry count for this address book."
                : $"The entry count for this address book could not be read: {ex.Message}";
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref entries);
        }

        return info;
    }

    /// <summary>
    /// Finds the requested address book, or the one Outlook opens first when none was named.
    /// Returns null when nothing matched, so the caller can refuse by name rather than quietly
    /// answering about a different book.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.AddressList? ResolveAddressList(Outlook.NameSpace session, string? requested)
    {
        Outlook.AddressLists? lists = null;

        try
        {
            lists = session.AddressLists;
            int count = SafeGetInt(() => lists.Count);

            Outlook.AddressList? fallback = null;

            try
            {
                for (int index = 1; index <= count; index++)
                {
                    Outlook.AddressList? list = null;
                    bool keep = false;

                    try
                    {
                        list = lists[index];

                        if (Matches(list, requested))
                        {
                            keep = true;
                            return list;
                        }

                        if (string.IsNullOrWhiteSpace(requested) && index == 1)
                        {
                            // The first book stands in for "no book is marked as the initial one",
                            // which some profiles are. It is held rather than released, and the
                            // outer finally lets it go again if a later book turns out to match.
                            keep = true;
                            fallback = list;
                        }
                    }
                    finally
                    {
                        if (!keep)
                        {
                            OutlookInteropRunner.ReleaseComObject(ref list);
                        }
                    }
                }

                Outlook.AddressList? matched = fallback;
                fallback = null;
                return matched;
            }
            finally
            {
                // Non-null only when an earlier iteration parked the first book and a later one
                // matched, in which case the parked book is nobody's to release but ours.
                OutlookInteropRunner.ReleaseComObject(ref fallback);
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref lists);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool Matches(Outlook.AddressList list, string? requested)
    {
        string type = DescribeAddressListType(list);

        if (string.IsNullOrWhiteSpace(requested))
        {
            return SafeGetBool(() => list.IsInitialAddressList);
        }

        string wanted = requested.Trim();

        if (wanted.Equals("gal", StringComparison.OrdinalIgnoreCase)
            || wanted.Equals("global", StringComparison.OrdinalIgnoreCase)
            || wanted.Equals("global address list", StringComparison.OrdinalIgnoreCase))
        {
            return type == "exchange-global-address-list";
        }

        if (wanted.Equals("contacts", StringComparison.OrdinalIgnoreCase)
            || wanted.Equals("contact", StringComparison.OrdinalIgnoreCase))
        {
            return type == "outlook-address-list";
        }

        string? name = SafeGet(() => list.Name);
        return name != null && name.Equals(wanted, StringComparison.OrdinalIgnoreCase);
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string DescribeAddressListType(Outlook.AddressList list)
    {
        try
        {
            return list.AddressListType switch
            {
                Outlook.OlAddressListType.olExchangeGlobalAddressList => "exchange-global-address-list",
                Outlook.OlAddressListType.olExchangeContainer => "exchange-container",
                Outlook.OlAddressListType.olOutlookAddressList => "outlook-address-list",
                Outlook.OlAddressListType.olOutlookLdapAddressList => "outlook-ldap-address-list",
                Outlook.OlAddressListType.olCustomAddressList => "custom-address-list",
                _ => "unknown"
            };
        }
        catch (COMException)
        {
            // AddressListType arrived in Outlook 2007. A book that will not say what it is gets
            // "unknown", never a number: olExchangeGlobalAddressList is 0, so a numeric projection
            // would make the Global Address List indistinguishable from an unset default.
            return "unknown";
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string DescribeEntryType(Outlook.OlAddressEntryUserType? userType) => userType switch
    {
        Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry => "exchange-user",
        Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry => "exchange-distribution-list",
        Outlook.OlAddressEntryUserType.olExchangePublicFolderAddressEntry => "exchange-public-folder",
        Outlook.OlAddressEntryUserType.olExchangeAgentAddressEntry => "exchange-agent",
        Outlook.OlAddressEntryUserType.olExchangeOrganizationAddressEntry => "exchange-organization",
        Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry => "exchange-remote-user",
        Outlook.OlAddressEntryUserType.olOutlookContactAddressEntry => "outlook-contact",
        Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry => "outlook-distribution-list",
        Outlook.OlAddressEntryUserType.olLdapAddressEntry => "ldap",
        Outlook.OlAddressEntryUserType.olSmtpAddressEntry => "smtp",
        Outlook.OlAddressEntryUserType.olOtherAddressEntry => "other",
        _ => "unknown"
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.OlAddressEntryUserType? SafeGetUserType(Outlook.AddressEntry entry)
    {
        try
        {
            return entry.AddressEntryUserType;
        }
        catch (COMException)
        {
            return null;
        }
    }

    // ── Small helpers ───────────────────────────────────────────────────────

    /// <summary>
    /// Splits an addressee argument on semicolons, and on semicolons only.
    ///
    /// <para>
    /// Commas are deliberately <b>not</b> separators. <c>Smith, Jane</c> is the canonical Exchange
    /// Global Address List display-name shape, so splitting on commas would take the single most
    /// common form of the exact input this action exists to resolve and turn it into two addressees,
    /// neither of which resolves. Outlook itself uses <c>;</c> between recipients for the same
    /// reason. Accepting commas as well would be friendlier in the abstract and wrong in the case
    /// that matters.
    /// </para>
    /// </summary>
    private static List<string> SplitRecipients(string? recipients)
    {
        if (string.IsNullOrWhiteSpace(recipients))
        {
            return [];
        }

        return recipients
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .ToList();
    }

    /// <summary>
    /// Whether a string is a mailable address rather than an X500 legacyExchangeDN. The DN check
    /// is the point: <c>/o=ExchangeLabs/ou=.../cn=Recipients/cn=abc</c> contains no '@' but a
    /// looser test would still let some provider-specific identifiers through.
    /// </summary>
    private static bool LooksLikeSmtpAddress([NotNullWhen(true)] string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (value.StartsWith("/o=", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        int at = value.IndexOf('@');
        return at > 0 && at < value.Length - 1;
    }

    private static string? NullIfBlank(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value;

    /// <summary>
    /// Reads a string property, recording an Object Model Guard denial against
    /// <paramref name="propertyName"/> rather than letting it look identical to "not present".
    /// See Rule 22.
    /// </summary>
    private static string? SafeGet(Func<string?> getter, string propertyName, List<string> accessDenied)
    {
        try
        {
            return getter();
        }
        catch (COMException ex) when (OutlookInteropRunner.IsObjectModelGuardDenial(ex))
        {
            RecordAccessDenied(accessDenied, propertyName);
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static void RecordAccessDenied(List<string> accessDenied, string memberName)
    {
        if (!accessDenied.Contains(memberName))
        {
            accessDenied.Add(memberName);
        }
    }

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

    private static bool SafeGetBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch (COMException)
        {
            return false;
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

    /// <summary>
    /// Reads a COM object property that may legitimately be unavailable - <c>GetExchangeUser</c>
    /// on a disconnected client, for instance.
    ///
    /// <para>
    /// An Object Model Guard denial is recorded against <paramref name="memberName"/> rather than
    /// being returned as a plain null. Every member this is used for -
    /// <c>Recipient.AddressEntry</c>, <c>GetExchangeUser</c>, <c>GetExchangeDistributionList</c>,
    /// <c>GetContact</c>, <c>PropertyAccessor</c> - is on Outlook's protected list, and this
    /// surface exists to validate an addressee before sending. "Outlook has no such person" and
    /// "Outlook refused to tell me" lead to opposite actions: correct the address, or treat the
    /// answer as unknown and do not call the send validated. Collapsing them would be a silent
    /// wrong answer in the one place the feature has to be trustworthy. See Rule 22.
    /// </para>
    /// </summary>
    private static T? SafeGetComObject<T>(Func<T?> getter, string memberName, List<string> accessDenied)
        where T : class
    {
        try
        {
            return getter();
        }
        catch (COMException ex) when (OutlookInteropRunner.IsObjectModelGuardDenial(ex))
        {
            RecordAccessDenied(accessDenied, memberName);
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }
}
