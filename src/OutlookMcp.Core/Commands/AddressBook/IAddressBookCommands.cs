using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.AddressBook;

[ServiceCategory("addressbook")]
[McpTool("addressbook", Title = "Outlook Address Book Operations", Destructive = false, Category = "addressbook",
    Description = "Look up addressees in Outlook's address books before using them, instead of "
    + "discovering at send time that a name did not resolve. Read-only: nothing here creates, "
    + "changes or sends anything. "
    + "Use resolve to check one or more addressees and get their real SMTP addresses. It accepts "
    + "display names, Exchange aliases and full email addresses, separated by semicolons, and "
    + "answers per addressee: allResolved is the flag to check before sending, and unresolvedNames "
    + "says which ones are wrong. A name Outlook cannot find is a success with resolved false - not "
    + "an error - because 'no such person' and 'Outlook could not be reached' are different answers. "
    + "An ambiguous name also comes back unresolved: Outlook offers no way to list the candidates, "
    + "so pass the full SMTP address to disambiguate. Note that resolution is a weak existence "
    + "check for anything SMTP-shaped: a syntactically valid address resolves as a one-off whether "
    + "or not the mailbox exists, and smtpAddressSource says which kind of answer you got. "
    + "smtpAddress is always a real email address. Outlook's own Address property returns an X500 "
    + "legacyExchangeDN such as '/o=ExchangeLabs/ou=.../cn=...' for an Exchange entry, which looks "
    + "like an address and cannot be mailed; that value is reported separately as rawAddress and is "
    + "never passed off as an email address. "
    + "Use list-address-lists to discover which books exist - the Exchange Global Address List, "
    + "Contacts folders, LDAP directories - before searching one. hasGlobalAddressList is false on "
    + "a profile with no Exchange account, where colleagues simply cannot be looked up. "
    + "Use list-entries to browse one book. Outlook's object model has no server-side address book "
    + "search, so this scans, and the answer is always 'what matched in the part that was examined'. "
    + "startsWith filters by display-name prefix client-side. Check scanLimitReached: when it is "
    + "true an empty result is not evidence that nobody matches, and resolve is the right call for a "
    + "person you can already name. "
    + "OBJECT MODEL GUARD: recipients and address entries are among the members Outlook protects "
    + "against out-of-process callers. Every action here can be refused by a modal security prompt "
    + "that no program can answer, in which case the call fails with an explanation rather than "
    + "hanging or returning a wrong answer. Individual properties that were refused while the call "
    + "otherwise succeeded are named in accessDenied, so a missing value is never confused with a "
    + "value the directory does not hold.")]
public interface IAddressBookCommands
{
    /// <summary>
    /// Lists the address books attached to the profile.
    ///
    /// <para>
    /// OBJECT MODEL GUARD: enumerating <c>NameSpace.AddressLists</c> is a protected operation and
    /// may be refused. <paramref name="includeEntryCount"/> additionally touches
    /// <c>AddressList.AddressEntries</c>, which is separately protected and, on a corporate Global
    /// Address List, expensive; it is off by default for both reasons.
    /// </para>
    /// </summary>
    /// <param name="includeEntryCount">Count the entries in each book. Off by default: it is a protected and potentially very slow read.</param>
    [ServiceAction("list-address-lists")]
    AddressListCollectionResult ListAddressLists(bool includeEntryCount = false);

    /// <summary>
    /// Resolves one or more addressees against the address book and reports their SMTP addresses.
    ///
    /// <para>
    /// This is the pre-send check. An addressee Outlook cannot find comes back with
    /// <c>resolved: false</c> on an otherwise successful call, so a caller can name the bad entry
    /// instead of sending and waiting for a bounce.
    /// </para>
    ///
    /// <para>
    /// OBJECT MODEL GUARD: every property and method on <c>Recipient</c> is protected, as are
    /// <c>AddressEntry.Address</c>, <c>AddressEntry.GetExchangeUser</c> and
    /// <c>ExchangeUser.PrimarySmtpAddress</c>. This action is therefore the most guard-exposed in
    /// the product. A denial fails the call with an explanation; a denial of one property while
    /// the rest succeeded is named in <c>accessDenied</c>.
    /// </para>
    /// </summary>
    /// <param name="recipients">One or more display names, Exchange aliases or email addresses, separated by semicolons or commas.</param>
    /// <param name="includeDetails">Also read job title, department, office and alias for Exchange users. Each is an extra directory read per addressee.</param>
    [ServiceAction("resolve")]
    AddressResolveResult Resolve(string recipients, bool includeDetails = true);

    /// <summary>
    /// Browses the entries in one address book.
    ///
    /// <para>
    /// Outlook exposes no server-side search over an address book - there is no <c>Restrict</c> or
    /// <c>Find</c> on <c>AddressEntries</c> - so this scans from the start of the book and stops at
    /// <paramref name="scanLimit"/>. <c>scanLimitReached</c> in the result says whether the book
    /// ran out or the scan did, because an empty answer means different things in the two cases.
    /// </para>
    ///
    /// <para>
    /// OBJECT MODEL GUARD: <c>AddressList.AddressEntries</c> and every cursor method on it
    /// (<c>GetFirst</c>, <c>GetNext</c>, <c>Item</c>) are protected members and may be refused.
    /// </para>
    /// </summary>
    /// <param name="addressList">Which book to read, by name, or the aliases 'gal' and 'contacts'. Defaults to the book Outlook opens first.</param>
    /// <param name="startsWith">Keep only entries whose display name begins with this, compared case-insensitively. Applied while scanning, since Outlook cannot filter server-side.</param>
    /// <param name="maxCount">How many entries to return, 1 to 100.</param>
    /// <param name="scanLimit">How many entries to examine before giving up, 1 to 5000. A corporate Global Address List is far larger than this.</param>
    /// <param name="includeSmtpAddress">Resolve each returned entry to an SMTP address. One extra directory read per returned entry.</param>
    [ServiceAction("list-entries")]
    AddressEntryListResult ListEntries(
        string? addressList = null,
        string? startsWith = null,
        int maxCount = 25,
        int scanLimit = 500,
        bool includeSmtpAddress = true);
}
