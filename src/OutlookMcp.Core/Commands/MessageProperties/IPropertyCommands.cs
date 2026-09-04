using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.MessageProperties;

[ServiceCategory("property")]
[McpTool("property", Title = "Outlook Message Property Operations", Destructive = false, Category = "property",
    Description = "Read the properties Outlook does not project onto its own objects: internet "
    + "message headers, named MAPI properties and custom user properties. Read-only - there is no "
    + "way to write a property through this tool. "
    + "Use get-headers for the internet headers of a received message: the delivery path in the "
    + "Received chain, SPF and DKIM verdicts in Authentication-Results and Received-SPF, "
    + "List-Unsubscribe, and the original sender behind a forward. Headers come back parsed into "
    + "names and values, duplicates preserved in transport order, with folded continuation lines "
    + "joined back on. A header block runs to tens of kilobytes, so pass headerName to get one "
    + "header rather than all of them, and set includeRaw only when the unparsed block is genuinely "
    + "wanted. "
    + "A DRAFT HAS NO HEADERS. Nothing composed locally ever traversed a transport, so it carries "
    + "none - the call succeeds with headersPresent false. That is an answer, not a failure, and it "
    + "is not evidence that anything went wrong. The same is true of an item delivered entirely "
    + "inside one organisation. "
    + "Use get-known for a curated set of properties that are commonly useful and easy to get wrong "
    + "by hand: the Internet message id, the in-reply-to id, the sender's and sent-on-behalf-of SMTP "
    + "addresses, and whether the message was replied to or forwarded. Every property comes back "
    + "with its own status whether or not the item carries it, so a missing entry can never be "
    + "mistaken for one that was not asked for. "
    + "Use get-property to read ANY MAPI property by its DASL name. This is deliberately not "
    + "restricted to a curated list: it will read anything the store holds on the item you name, "
    + "including properties the rest of this surface deliberately does not project. It is read-only "
    + "and it cannot reach an item you could not already open with mail.read, but it is not a "
    + "narrow, audited window onto the item either - treat it accordingly and prefer get-headers or "
    + "get-known when they answer the question. "
    + "Use list-user-properties for custom properties a form, an add-in or a user put on an item. "
    + "Most items have none, and an empty list is a success. "
    + "A property the item does not carry is reported as not-present on a successful call, never as "
    + "an error, because absence is the single most common outcome for a MAPI property. Some stores "
    + "return an empty value instead of reporting a property missing, which is reported as status "
    + "empty; found is false for both, so 'is there a usable value here' is one check. "
    + "OBJECT MODEL GUARD: PropertyAccessor is itself one of the members Outlook protects against "
    + "out-of-process callers, so any action here can be refused by a modal security prompt that no "
    + "program can answer. A refused property is reported as blocked, which is a different status "
    + "from not-present: blocked means the value exists and was withheld.")]
public interface IPropertyCommands
{
    /// <summary>
    /// Reads the internet message headers of an item (<c>PR_TRANSPORT_MESSAGE_HEADERS</c>).
    ///
    /// <para>
    /// Absence is a normal outcome. An item that never traversed an SMTP transport carries no
    /// headers, so a draft, a locally created item, and often an item delivered inside a single
    /// organisation all report <c>headersPresent: false</c> on a successful call.
    /// </para>
    ///
    /// <para>
    /// OBJECT MODEL GUARD: <c>PropertyAccessor</c> is a protected member on every Outlook item
    /// type, so this read may be refused by a security prompt that cannot be answered
    /// programmatically.
    /// </para>
    /// </summary>
    /// <param name="entryId">The item to read. Falls back to the item open or selected in Outlook when omitted.</param>
    /// <param name="storeId">The store the entry id belongs to. An entry id is only meaningful together with its store.</param>
    /// <param name="useActiveMail">Fall back to the item currently open or selected in Outlook when no entry id is given.</param>
    /// <param name="headerName">Return only headers with this name, compared case-insensitively. A whole header block is often tens of kilobytes.</param>
    /// <param name="includeRaw">Also return the unparsed header block. Off by default because of its size.</param>
    [ServiceAction("get-headers")]
    MessageHeaderResult GetHeaders(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? headerName = null,
        bool includeRaw = false);

    /// <summary>
    /// Reads a curated set of MAPI properties that are commonly useful and easy to get wrong by
    /// hand: the wrong property type suffix simply fails, and the tags are not guessable.
    ///
    /// <para>
    /// Every property appears in the result with its own status whether or not the item carries it.
    /// Absence is expected and specific: <c>PR_SENT_REPRESENTING_SMTP_ADDRESS</c> is documented as
    /// being left unset on the local copy of a sent message, for instance.
    /// </para>
    ///
    /// <para>
    /// OBJECT MODEL GUARD: as for <see cref="GetHeaders"/>, obtaining the <c>PropertyAccessor</c>
    /// is a protected operation.
    /// </para>
    /// </summary>
    /// <param name="entryId">The item to read. Falls back to the item open or selected in Outlook when omitted.</param>
    /// <param name="storeId">The store the entry id belongs to.</param>
    /// <param name="useActiveMail">Fall back to the item currently open or selected in Outlook when no entry id is given.</param>
    [ServiceAction("get-known")]
    MessagePropertySetResult GetKnown(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    /// <summary>
    /// Reads one MAPI property by its DASL name.
    ///
    /// <para>
    /// This is deliberately unrestricted: any property the store holds on the named item can be
    /// read, including properties the rest of this surface does not project. It is read-only, and
    /// it cannot reach an item the caller could not already open, but it is not a narrow window
    /// onto the item either. Prefer <see cref="GetHeaders"/> or <see cref="GetKnown"/> where they
    /// answer the question.
    /// </para>
    ///
    /// <para>
    /// A property the item does not carry is reported as <c>not-present</c> on a successful call.
    /// Outlook signals that with <c>MAPI_E_NOT_FOUND</c>, which is an ordinary absence rather than
    /// a failure.
    /// </para>
    ///
    /// <para>
    /// OBJECT MODEL GUARD: obtaining the <c>PropertyAccessor</c> is protected. A refused read is
    /// reported as <c>blocked</c>, which is not the same as <c>not-present</c>.
    /// </para>
    /// </summary>
    /// <param name="dasl">The full DASL property name, e.g. 'http://schemas.microsoft.com/mapi/proptag/0x007D001F'. Property references are case-sensitive.</param>
    /// <param name="entryId">The item to read. Falls back to the item open or selected in Outlook when omitted.</param>
    /// <param name="storeId">The store the entry id belongs to.</param>
    /// <param name="useActiveMail">Fall back to the item currently open or selected in Outlook when no entry id is given.</param>
    [ServiceAction("get-property")]
    MessagePropertyResult GetProperty(
        string dasl,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    /// <summary>
    /// Lists the custom user properties on an item - the ones a form, an add-in or a user added.
    /// Most items have none, and an empty list is a successful answer.
    ///
    /// <para>
    /// OBJECT MODEL GUARD: reading a user property whose name matches a built-in address-bearing
    /// property is protected; ordinary custom properties are not.
    /// </para>
    /// </summary>
    /// <param name="entryId">The item to read. Falls back to the item open or selected in Outlook when omitted.</param>
    /// <param name="storeId">The store the entry id belongs to.</param>
    /// <param name="useActiveMail">Fall back to the item currently open or selected in Outlook when no entry id is given.</param>
    [ServiceAction("list-user-properties")]
    UserPropertyListResult ListUserProperties(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);
}
