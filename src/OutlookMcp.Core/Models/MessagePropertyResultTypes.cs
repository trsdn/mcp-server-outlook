using System.Text.Json.Serialization;

namespace OutlookMcp.Core.Models;

// ── Message properties: internet headers, MAPI properties, user properties (#15) ──

/// <summary>
/// The internet message headers an item carries, parsed. Absence is a normal outcome and is
/// reported as a success: an item that never traversed an SMTP transport - a draft, an item created
/// locally, an internal calendar response - simply has none.
/// </summary>
public class MessageHeaderResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    /// <summary>The Outlook item class this was read from, e.g. <c>MailItem</c>.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    /// <summary>
    /// False when the item carries no transport headers. This is the ordinary answer for anything
    /// that never left the client, and it is not an error.
    /// </summary>
    public bool HeadersPresent { get; set; }

    /// <summary>
    /// Headers in the order the transport wrote them, duplicates preserved. <c>Received</c> appears
    /// once per hop and the order is the delivery path in reverse, so collapsing duplicates would
    /// destroy the only thing that block is good for.
    /// </summary>
    public List<MessageHeaderInfo> Headers { get; set; } = [];

    /// <summary>Equal to the <see cref="Headers"/> count, after any name filter.</summary>
    public int HeaderCount { get; set; }

    /// <summary>The unparsed header block, and only when <c>includeRaw</c> was set.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Raw { get; set; }

    /// <summary>
    /// Characters in the raw header block, reported whether or not <see cref="Raw"/> was asked
    /// for. A header block on a message that crossed several relays runs to tens of kilobytes.
    /// </summary>
    public int RawLength { get; set; }

    /// <summary>Why the headers are absent, or which filter narrowed them.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

public class MessageHeaderInfo
{
    /// <summary>The header name without its colon, e.g. <c>Authentication-Results</c>.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// The value with continuation lines folded back in, so a long header arrives whole rather
    /// than split across phantom entries.
    /// </summary>
    public string Value { get; set; } = string.Empty;
}

/// <summary>
/// A curated set of MAPI properties that are commonly worth reading and easy to get wrong by hand.
/// Every requested property appears in the result whether or not the item carries it, each with its
/// own status, so a missing entry can never be mistaken for one that was not asked for.
/// </summary>
public class MessagePropertySetResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    public List<MessagePropertyInfo> Properties { get; set; } = [];
}

public class MessagePropertyInfo
{
    /// <summary>A readable name for the property, e.g. <c>internetMessageId</c>.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>The DASL property name this was read from, so the read can be repeated by hand.</summary>
    public string Dasl { get; set; } = string.Empty;

    /// <summary>False when the item does not carry this property, or the read was refused.</summary>
    public bool Found { get; set; }

    /// <summary>
    /// <c>ok</c>, <c>empty</c>, <c>not-present</c>, <c>blocked</c>, <c>unsupported-or-blocked</c>
    /// or <c>error</c>. Never null: a property with no status would be indistinguishable from one
    /// that was never attempted.
    ///
    /// <para>
    /// <c>empty</c> and <c>not-present</c> both mean there is no usable value, and
    /// <see cref="Found"/> is false for both. They are kept apart because Exchange returns an empty
    /// string for some tags where MAPI would report the property missing, and reporting that as a
    /// found value would answer "yes, this message has one" while handing back nothing.
    /// </para>
    /// </summary>
    public string Status { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; set; }

    /// <summary>A human-readable reading of a coded value, where one exists.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Meaning { get; set; }

    /// <summary>Why the property is absent, when the reason is worth stating.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

/// <summary>One arbitrary MAPI property, read by its DASL name.</summary>
public class MessagePropertyResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    /// <summary>The DASL property name that was read.</summary>
    public string Dasl { get; set; } = string.Empty;

    public bool Found { get; set; }

    /// <summary>
    /// <c>ok</c>, <c>empty</c>, <c>not-present</c>, <c>blocked</c>, <c>unsupported-or-blocked</c>
    /// or <c>error</c>. <see cref="Found"/> is true only for <c>ok</c>.
    /// </summary>
    public string Status { get; set; } = string.Empty;

    /// <summary>
    /// The value rendered as text. For a binary property this is the hex form, never
    /// <c>System.Byte[]</c>.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; set; }

    /// <summary>
    /// <c>string</c>, <c>int</c>, <c>bool</c>, <c>dateTime</c>, <c>binary</c>, <c>array</c> or
    /// <c>unknown</c>. MAPI property types are not interchangeable and a caller has to know which
    /// one it got.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ValueType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? BinaryLength { get; set; }

    /// <summary>A binary value in base64, for a caller that wants the bytes.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Base64 { get; set; }

    /// <summary>
    /// A binary value in the hex form Outlook itself uses. This is the form an entry id has to be
    /// in to be passed back to Outlook, so it is reported alongside the base64.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Hex { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

/// <summary>Custom properties a form, an add-in or a user put on an item.</summary>
public class UserPropertyListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    public List<UserPropertyInfo> UserProperties { get; set; } = [];

    /// <summary>Equal to the <see cref="UserProperties"/> count. Zero is the common case.</summary>
    public int Count { get; set; }
}

public class UserPropertyInfo
{
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Outlook's own type for the property as a name: <c>text</c>, <c>number</c>,
    /// <c>date-time</c>, <c>yes-no</c>, <c>keywords</c>, and so on.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Type { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; set; }

    /// <summary>The CLR shape the value arrived as: <c>string</c>, <c>int</c>, and so on.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ValueType { get; set; }

    /// <summary>Why the value is absent when the property itself was readable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}
