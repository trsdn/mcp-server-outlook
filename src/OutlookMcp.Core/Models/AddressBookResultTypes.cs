using System.Text.Json.Serialization;

namespace OutlookMcp.Core.Models;

// ── Address book / GAL (#15) ────────────────────────────────────────────────
//
// Kept out of ResultTypes.cs deliberately. That file is a 1900-line grab bag inherited from an
// earlier product; new domains get their own file so the address-book surface can be read in one
// sitting.

/// <summary>
/// The address books attached to the profile: the Exchange Global Address List, Contacts folders
/// exposed as address lists, LDAP directories and any custom lists.
/// </summary>
public class AddressListCollectionResult : ResultBase
{
    public List<AddressListInfo> AddressLists { get; set; } = [];

    /// <summary>Equal to the <see cref="AddressLists"/> count. Present so a caller need not count.</summary>
    public int Count { get; set; }

    /// <summary>
    /// True when at least one list is an Exchange Global Address List. False on a profile with no
    /// Exchange account, where directory lookup of colleagues is simply not available and the only
    /// addressees that exist are the ones in local Contacts.
    /// </summary>
    public bool HasGlobalAddressList { get; set; }
}

public class AddressListInfo
{
    /// <summary>The name to pass back as <c>addressList</c> to enumerate this book.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>1-based position in Outlook's own resolution order.</summary>
    public int Index { get; set; }

    /// <summary>
    /// Outlook's classification as a name: <c>exchange-global-address-list</c>,
    /// <c>exchange-container</c>, <c>outlook-address-list</c>, <c>outlook-ldap-address-list</c> or
    /// <c>custom-address-list</c>. Never a raw enum number - the Global Address List is 0, so a
    /// number would be indistinguishable from an unset default.
    /// </summary>
    public string AddressListType { get; set; } = string.Empty;

    public bool IsReadOnly { get; set; }

    /// <summary>
    /// True for the book Outlook opens first in its own Select Names dialog. This is not the same
    /// as "is the Global Address List" - check <see cref="AddressListType"/> for that.
    /// </summary>
    public bool IsInitialAddressList { get; set; }

    /// <summary>
    /// Entries in this book, and only when <c>includeEntryCount</c> was set. Counting means
    /// touching <c>AddressList.AddressEntries</c>, which is Object Model Guard protected and, on a
    /// corporate Global Address List, expensive.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? EntryCount { get; set; }

    /// <summary>Why <see cref="EntryCount"/> is absent although it was asked for.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

/// <summary>
/// The outcome of resolving one or more addressees against the address book. A name that does not
/// resolve is reported here rather than as a failure: the lookup worked, and the answer is "no
/// such addressee".
/// </summary>
public class AddressResolveResult : ResultBase
{
    public List<ResolvedRecipientInfo> Recipients { get; set; } = [];

    public int RequestedCount { get; set; }

    public int ResolvedCount { get; set; }

    /// <summary>
    /// True only when every requested addressee resolved. This is the flag to check before
    /// sending anything.
    /// </summary>
    public bool AllResolved { get; set; }

    /// <summary>
    /// The queries that did not resolve, in the order they were given. A convenience mirror of
    /// <see cref="Recipients"/>, so "which ones are wrong" costs no filtering.
    /// </summary>
    public List<string> UnresolvedNames { get; set; } = [];
}

public class ResolvedRecipientInfo
{
    /// <summary>The string that was looked up, exactly as supplied.</summary>
    public string Query { get; set; } = string.Empty;

    /// <summary>
    /// False when Outlook found no match, and also when the name was ambiguous. Outlook's object
    /// model does not distinguish the two and offers no way to list the ambiguous candidates, so
    /// an ambiguous name is reported as unresolved; supply the full SMTP address to disambiguate.
    /// </summary>
    public bool Resolved { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DisplayName { get; set; }

    /// <summary>
    /// The deliverable email address. Null when the entry has none that could be read - and never
    /// the X500 legacyExchangeDN, which is what Outlook's own <c>Address</c> property returns for
    /// an Exchange entry and which is not an email address.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SmtpAddress { get; set; }

    /// <summary>
    /// How <see cref="SmtpAddress"/> was obtained: <c>exchange-user</c>,
    /// <c>exchange-distribution-list</c>, <c>smtp-entry</c>, <c>contact</c> or
    /// <c>property-accessor</c>. Present so a caller can tell a directory-backed answer from a
    /// one-off SMTP string that was never checked against anything.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SmtpAddressSource { get; set; }

    /// <summary>
    /// The provider's own address string. For an Exchange entry this is the X500
    /// legacyExchangeDN (<c>/o=.../cn=...</c>), which is useful for diagnosis and useless for
    /// addressing mail. Reported alongside <see cref="SmtpAddress"/>, never instead of it.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RawAddress { get; set; }

    /// <summary>The MAPI address type, conventionally <c>EX</c> for Exchange or <c>SMTP</c>.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AddressType { get; set; }

    /// <summary>
    /// What kind of directory object this is, as a name: <c>exchange-user</c>,
    /// <c>exchange-distribution-list</c>, <c>exchange-remote-user</c>, <c>outlook-contact</c>,
    /// <c>smtp</c>, and so on.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryType { get; set; }

    /// <summary>
    /// True for a group rather than a person. Sending to a group is not the same act as sending to
    /// one colleague, so this is surfaced rather than left to be inferred from a type name.
    /// </summary>
    public bool IsDistributionList { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Alias { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? JobTitle { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Department { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? OfficeLocation { get; set; }

    /// <summary>
    /// Properties Outlook's Object Model Guard refused, by name. An empty list means nothing was
    /// blocked; a name here means the value is missing because it was denied, not because the
    /// directory does not hold it.
    /// </summary>
    public List<string> AccessDenied { get; set; } = [];

    /// <summary>Why an otherwise resolved entry carries no SMTP address.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

/// <summary>
/// A page of entries from one address book. The counts are the load-bearing part: an address book
/// has no server-side search in the Outlook object model, so the answer is always "what was found
/// in the part that was examined", never "everyone who matches".
/// </summary>
public class AddressEntryListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AddressListName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AddressListType { get; set; }

    public List<AddressBookEntryInfo> Entries { get; set; } = [];

    /// <summary>Equal to the <see cref="Entries"/> count.</summary>
    public int ReturnedCount { get; set; }

    /// <summary>Entries actually examined, which exceeds the returned count when filtering.</summary>
    public int ScannedCount { get; set; }

    /// <summary>True when the book holds more entries than were examined.</summary>
    public bool Truncated { get; set; }

    /// <summary>
    /// True when the scan stopped at <c>scanLimit</c> rather than at the end of the book. A prefix
    /// filter matching nothing in the first slice of a large Global Address List reports this, and
    /// an empty result with this set is not evidence that nobody matches.
    /// </summary>
    public bool ScanLimitReached { get; set; }
}

public class AddressBookEntryInfo
{
    public string DisplayName { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SmtpAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SmtpAddressSource { get; set; }

    /// <summary>The provider address; an X500 legacyExchangeDN for an Exchange entry.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RawAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AddressType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryType { get; set; }

    public bool IsDistributionList { get; set; }
}
