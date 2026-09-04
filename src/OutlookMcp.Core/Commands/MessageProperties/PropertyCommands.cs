using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.MessageProperties;

/// <summary>
/// Reads internet message headers, named MAPI properties and user properties through
/// <c>PropertyAccessor</c> (#15).
///
/// <para>
/// Three behaviours drive the shape of everything here, and all three were measured rather than
/// assumed. <c>GetProperty</c> <b>throws</b> when a property is absent instead of returning null,
/// and the HRESULT that means "simply not there" (<c>MAPI_E_NOT_FOUND</c>, 0x8004010F) is a
/// different one from the HRESULT that means "withheld" (<c>MAPI_E_NOT_SUPPORTED</c> 0x80040102 or
/// <c>E_ACCESSDENIED</c> 0x80070005). Collapsing those into one "no value" answer would tell a
/// caller that a message has no sender address when in fact Outlook refused to say.
/// </para>
///
/// <para>
/// The second is that a binary property arrives as a byte array, so anything that stringifies it
/// naively emits the literal <c>System.Byte[]</c> and reports success. The third is that a draft
/// has no transport headers at all, which is the reason absence is modelled as an outcome rather
/// than an error throughout.
/// </para>
///
/// <para>
/// <b>On arbitrary DASL reads.</b> <see cref="GetProperty"/> deliberately accepts any DASL property
/// name rather than a curated allow-list. It is read-only, and it cannot reach an item the caller
/// could not already open in full through <c>mail.read</c> - so it grants no access that the rest
/// of this surface does not already grant. What it does do is expose properties this surface
/// deliberately does not project, and that is stated plainly in the tool description rather than
/// left to be discovered. The alternative, a fixed list, would be permanently incomplete over a
/// property space with thousands of members, and the curation would be guesswork.
/// </para>
/// </summary>
public class PropertyCommands : IPropertyCommands
{
    private const string PrTransportMessageHeadersUnicode = "http://schemas.microsoft.com/mapi/proptag/0x007D001F";
    private const string PrTransportMessageHeadersAnsi = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

    private const int MapiNotFound = unchecked((int)0x8004010F);
    private const int MapiNotSupported = unchecked((int)0x80040102);
    private const int AccessDenied = unchecked((int)0x80070005);
    private const int Abort = unchecked((int)0x80004004);

    /// <summary>
    /// The namespaces <c>PropertyAccessor</c> understands. A string that is not one of these is
    /// refused here rather than handed to Outlook, which answers "The property name is invalid" and
    /// says nothing about what a valid one looks like.
    /// </summary>
    private static readonly string[] DaslPrefixes =
    [
        "http://schemas.microsoft.com/mapi/",
        "https://schemas.microsoft.com/mapi/",
        "http://schemas.microsoft.com/exchange/",
        "https://schemas.microsoft.com/exchange/",
        "urn:schemas:",
        "urn:schemas-microsoft-com:",
        "DAV:"
    ];

    /// <summary>
    /// Properties worth reading and awkward to get right by hand: the tags are not guessable and
    /// the wrong type suffix simply fails.
    /// </summary>
    private static readonly (string Name, string Dasl)[] KnownProperties =
    [
        ("internetMessageId", "http://schemas.microsoft.com/mapi/proptag/0x1035001F"),
        ("inReplyToId", "http://schemas.microsoft.com/mapi/proptag/0x1042001F"),
        ("senderSmtpAddress", "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"),
        ("sentRepresentingSmtpAddress", "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"),
        ("lastVerbExecuted", "http://schemas.microsoft.com/mapi/proptag/0x10810003"),
        ("messageFlags", "http://schemas.microsoft.com/mapi/proptag/0x0E070003")
    ];

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MessageHeaderResult GetHeaders(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? headerName = null,
        bool includeRaw = false)
    {
        string? filter = NullIfBlank(headerName);

        return OutlookInteropRunner.Execute(
            "OutlookPropertyGetHeaders",
            (application, session) =>
            {
                var scope = new ItemScope();

                try
                {
                    object? item = scope.Resolve(application, session, entryId, storeId, useActiveMail);

                    if (item == null)
                    {
                        return new MessageHeaderResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedItemMessage(entryId)
                        };
                    }

                    var result = new MessageHeaderResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => GetEntryId(item)),
                        StoreId = NullIfBlank(storeId),
                        Subject = SafeGet(() => GetSubject(item)),
                        ItemType = DescribeItemType(item)
                    };

                    Outlook.PropertyAccessor? accessor = null;

                    try
                    {
                        accessor = GetPropertyAccessor(item);

                        if (accessor == null)
                        {
                            result.Note = $"An item of type {result.ItemType} exposes no property "
                                + "accessor, so its headers cannot be read.";
                            return result;
                        }

                        // Unicode first, then ANSI. Outlook coerces between the two string forms,
                        // but a store that holds only one refuses the other with MAPI_E_NOT_FOUND.
                        var read = ReadProperty(accessor, PrTransportMessageHeadersUnicode);

                        if (read.Status != "ok")
                        {
                            read = ReadProperty(accessor, PrTransportMessageHeadersAnsi);
                        }

                        string? raw = read.Value as string;

                        if (read.Status != "ok" || string.IsNullOrWhiteSpace(raw))
                        {
                            result.HeadersPresent = false;
                            result.Note = read.Status switch
                            {
                                "blocked" or "unsupported-or-blocked" =>
                                    "Outlook's security prompt refused the transport headers on this item.",
                                _ =>
                                    "This item carries no internet message headers. Anything composed "
                                    + "locally - a draft, a saved copy, an item delivered entirely inside "
                                    + "one organisation - never traversed an SMTP transport and so has none. "
                                    + "This is not evidence that the read failed."
                            };
                            return result;
                        }

                        result.HeadersPresent = true;
                        result.RawLength = raw.Length;
                        result.Headers = ParseHeaders(raw, filter);
                        result.HeaderCount = result.Headers.Count;

                        if (includeRaw)
                        {
                            result.Raw = raw;
                        }

                        if (filter != null)
                        {
                            result.Note = result.Headers.Count == 0
                                ? $"This item carries headers, but none called '{filter}'."
                                : $"Only headers called '{filter}' are shown; the item carries more.";
                        }

                        return result;
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref accessor);
                    }
                }
                finally
                {
                    scope.Release();
                }
            },
            ex => new MessageHeaderResult
            {
                Success = false,
                ErrorMessage = $"Failed to read the internet message headers: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MessagePropertySetResult GetKnown(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookPropertyGetKnown",
            (application, session) =>
            {
                var scope = new ItemScope();

                try
                {
                    object? item = scope.Resolve(application, session, entryId, storeId, useActiveMail);

                    if (item == null)
                    {
                        return new MessagePropertySetResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedItemMessage(entryId)
                        };
                    }

                    var result = new MessagePropertySetResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => GetEntryId(item)),
                        StoreId = NullIfBlank(storeId),
                        Subject = SafeGet(() => GetSubject(item)),
                        ItemType = DescribeItemType(item)
                    };

                    Outlook.PropertyAccessor? accessor = null;

                    try
                    {
                        accessor = GetPropertyAccessor(item);

                        if (accessor == null)
                        {
                            return new MessagePropertySetResult
                            {
                                Success = false,
                                ErrorMessage = $"An item of type {result.ItemType} exposes no property accessor."
                            };
                        }

                        foreach ((string name, string dasl) in KnownProperties)
                        {
                            var read = ReadProperty(accessor, dasl);

                            var info = new MessagePropertyInfo
                            {
                                Name = name,
                                Dasl = dasl,
                                Status = read.Status,
                                Found = read.Status == "ok"
                            };

                            if (info.Found)
                            {
                                var rendered = RenderValue(accessor, read.Value);
                                info.Value = rendered.Text;
                                info.Meaning = DescribeMeaning(name, read.Value);
                            }

                            info.Note = DescribeAbsence(name, info.Status, info.Value);
                            result.Properties.Add(info);
                        }

                        return result;
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref accessor);
                    }
                }
                finally
                {
                    scope.Release();
                }
            },
            ex => new MessagePropertySetResult
            {
                Success = false,
                ErrorMessage = $"Failed to read the item's MAPI properties: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MessagePropertyResult GetProperty(
        string dasl,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        string? requested = NullIfBlank(dasl);

        if (requested == null)
        {
            return new MessagePropertyResult
            {
                Success = false,
                ErrorMessage = "dasl is required for property.get-property: pass a full DASL property "
                    + "name such as 'http://schemas.microsoft.com/mapi/proptag/0x007D001F'."
            };
        }

        if (!DaslPrefixes.Any(prefix => requested.StartsWith(prefix, StringComparison.Ordinal)))
        {
            return new MessagePropertyResult
            {
                Success = false,
                Dasl = requested,
                ErrorMessage = $"'{requested}' is not a DASL property name. Use the full form, for "
                    + "example 'http://schemas.microsoft.com/mapi/proptag/0x007D001F' for a MAPI "
                    + "property tag, 'http://schemas.microsoft.com/mapi/string/{GUID}/name' for a "
                    + "named property, or 'urn:schemas:httpmail:subject'. Property references are "
                    + "case-sensitive."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookPropertyGetProperty",
            (application, session) =>
            {
                var scope = new ItemScope();

                try
                {
                    object? item = scope.Resolve(application, session, entryId, storeId, useActiveMail);

                    if (item == null)
                    {
                        return new MessagePropertyResult
                        {
                            Success = false,
                            Dasl = requested,
                            ErrorMessage = BuildUnresolvedItemMessage(entryId)
                        };
                    }

                    var result = new MessagePropertyResult
                    {
                        Success = true,
                        Dasl = requested,
                        EntryId = SafeGet(() => GetEntryId(item)),
                        StoreId = NullIfBlank(storeId),
                        Subject = SafeGet(() => GetSubject(item)),
                        ItemType = DescribeItemType(item)
                    };

                    Outlook.PropertyAccessor? accessor = null;

                    try
                    {
                        accessor = GetPropertyAccessor(item);

                        if (accessor == null)
                        {
                            return new MessagePropertyResult
                            {
                                Success = false,
                                Dasl = requested,
                                ErrorMessage = $"An item of type {result.ItemType} exposes no property accessor."
                            };
                        }

                        var read = ReadProperty(accessor, requested);
                        result.Status = read.Status;
                        result.Found = read.Status == "ok";

                        if (!result.Found)
                        {
                            result.Note = DescribeStatus(read.Status, read.Message);
                            return result;
                        }

                        var rendered = RenderValue(accessor, read.Value);
                        result.Value = rendered.Text;
                        result.ValueType = rendered.ValueType;
                        result.BinaryLength = rendered.BinaryLength;
                        result.Base64 = rendered.Base64;
                        result.Hex = rendered.Hex;
                        return result;
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref accessor);
                    }
                }
                finally
                {
                    scope.Release();
                }
            },
            ex => new MessagePropertyResult
            {
                Success = false,
                Dasl = requested,
                ErrorMessage = $"Failed to read '{requested}': {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public UserPropertyListResult ListUserProperties(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookPropertyListUserProperties",
            (application, session) =>
            {
                var scope = new ItemScope();

                try
                {
                    object? item = scope.Resolve(application, session, entryId, storeId, useActiveMail);

                    if (item == null)
                    {
                        return new UserPropertyListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnresolvedItemMessage(entryId)
                        };
                    }

                    var result = new UserPropertyListResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => GetEntryId(item)),
                        StoreId = NullIfBlank(storeId),
                        Subject = SafeGet(() => GetSubject(item)),
                        ItemType = DescribeItemType(item)
                    };

                    Outlook.UserProperties? properties = null;

                    try
                    {
                        properties = GetUserProperties(item);

                        if (properties == null)
                        {
                            return new UserPropertyListResult
                            {
                                Success = false,
                                ErrorMessage = $"An item of type {result.ItemType} exposes no user properties collection."
                            };
                        }

                        int count = SafeGetInt(() => properties.Count);

                        for (int index = 1; index <= count; index++)
                        {
                            Outlook.UserProperty? property = null;

                            try
                            {
                                property = properties[index];
                                result.UserProperties.Add(DescribeUserProperty(property, index));
                            }
                            finally
                            {
                                OutlookInteropRunner.ReleaseComObject(ref property);
                            }
                        }

                        result.Count = result.UserProperties.Count;
                        return result;
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref properties);
                    }
                }
                finally
                {
                    scope.Release();
                }
            },
            ex => new UserPropertyListResult
            {
                Success = false,
                ErrorMessage = $"Failed to read the item's user properties: {ex.Message}"
            });
    }

    // ── Property reading ────────────────────────────────────────────────────

    /// <summary>
    /// Reads one property and classifies the outcome.
    ///
    /// <para>
    /// The three-way split is the point. <c>MAPI_E_NOT_FOUND</c> means the item does not carry the
    /// property, which is ordinary. <c>MAPI_E_NOT_SUPPORTED</c> and <c>E_ACCESSDENIED</c> mean the
    /// value was withheld, which is not. Reporting both as "no value" would let a security refusal
    /// masquerade as an absent property - exactly the ambiguity Rule 22 exists to prevent.
    /// </para>
    ///
    /// <para>
    /// <c>MAPI_E_NOT_SUPPORTED</c> is genuinely ambiguous: Outlook also returns it for a property
    /// type <c>PropertyAccessor</c> cannot handle at all, such as <c>PT_OBJECT</c>. It is therefore
    /// reported as <c>unsupported-or-blocked</c> rather than asserted to be one or the other.
    /// </para>
    ///
    /// <para>
    /// There is a fourth outcome, and it was measured rather than expected. Asked for
    /// <c>PR_TRANSPORT_MESSAGE_HEADERS</c> or <c>PR_INTERNET_MESSAGE_ID</c> on a draft, an Exchange
    /// store does not raise <c>MAPI_E_NOT_FOUND</c> - it returns an <b>empty string</b>. Passing
    /// that through as a found value would answer "yes, this message has an Internet message id"
    /// and hand back nothing, so an empty string is reported as <c>empty</c> and never as
    /// <c>ok</c>. <c>found</c> is false for both <c>empty</c> and <c>not-present</c>, so a caller
    /// that only wants to know whether there is a usable value can ignore the distinction.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static PropertyRead ReadProperty(Outlook.PropertyAccessor accessor, string dasl)
    {
        try
        {
            object? value = accessor.GetProperty(dasl);

            if (value is string text && string.IsNullOrWhiteSpace(text))
            {
                return new PropertyRead("empty", null, null);
            }

            return new PropertyRead("ok", value, null);
        }
        catch (COMException ex)
        {
            string status = ex.HResult switch
            {
                MapiNotFound => "not-present",
                MapiNotSupported => "unsupported-or-blocked",
                AccessDenied => "blocked",
                Abort => "blocked",
                _ => "error"
            };

            return new PropertyRead(status, null, ex.Message);
        }
    }

    private readonly record struct PropertyRead(string Status, object? Value, string? Message);

    private readonly record struct RenderedValue(
        string? Text,
        string ValueType,
        int? BinaryLength,
        string? Base64,
        string? Hex);

    /// <summary>
    /// Projects a MAPI value onto something a JSON caller can use. A binary value is the case that
    /// matters: it arrives as a byte array, and anything that calls <c>ToString()</c> on it emits
    /// the literal "System.Byte[]" and reports success.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static RenderedValue RenderValue(Outlook.PropertyAccessor accessor, object? value)
    {
        switch (value)
        {
            case null:
                return new RenderedValue(null, "unknown", null, null, null);

            case string text:
                return new RenderedValue(text, "string", null, null, null);

            case bool flag:
                return new RenderedValue(flag ? "true" : "false", "bool", null, null, null);

            case int number:
                return new RenderedValue(number.ToString(CultureInfo.InvariantCulture), "int", null, null, null);

            case long number:
                return new RenderedValue(number.ToString(CultureInfo.InvariantCulture), "int", null, null, null);

            case double number:
                return new RenderedValue(number.ToString(CultureInfo.InvariantCulture), "number", null, null, null);

            case DateTime moment:
                // PT_SYSTIME arrives without time zone conversion and is not guaranteed to be UTC,
                // so it is rendered exactly as Outlook handed it over rather than being adjusted.
                return new RenderedValue(
                    moment.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
                    "dateTime", null, null, null);

            case byte[] bytes:
                string? hex = SafeGet(() => accessor.BinaryToString(bytes)) ?? Convert.ToHexString(bytes);
                return new RenderedValue(hex, "binary", bytes.Length, Convert.ToBase64String(bytes), hex);

            case Array array:
                var parts = new List<string>();

                foreach (object? element in array)
                {
                    parts.Add(element?.ToString() ?? string.Empty);
                }

                return new RenderedValue(string.Join("; ", parts), "array", null, null, null);

            default:
                return new RenderedValue(
                    value.ToString(), value.GetType().Name, null, null, null);
        }
    }

    // ── Header parsing ──────────────────────────────────────────────────────

    /// <summary>
    /// Parses an RFC 5322 header block, unfolding continuation lines.
    ///
    /// <para>
    /// A line beginning with whitespace continues the header above it. Splitting the block line by
    /// line without folding those back in produces entries with no name and silently truncates
    /// exactly the headers an agent cares about most - <c>Authentication-Results</c> and
    /// <c>Received</c> are almost always folded.
    /// </para>
    ///
    /// <para>
    /// Duplicates are preserved in order: a message carries one <c>Received</c> header per relay
    /// hop, and their order is the delivery path in reverse.
    /// </para>
    /// </summary>
    private static List<MessageHeaderInfo> ParseHeaders(string raw, string? nameFilter)
    {
        var headers = new List<MessageHeaderInfo>();
        string? currentName = null;
        var currentValue = new StringBuilder();

        void Flush()
        {
            if (currentName == null)
            {
                return;
            }

            if (nameFilter == null || currentName.Equals(nameFilter, StringComparison.OrdinalIgnoreCase))
            {
                headers.Add(new MessageHeaderInfo
                {
                    Name = currentName,
                    Value = currentValue.ToString().Trim()
                });
            }

            currentName = null;
            currentValue.Clear();
        }

        foreach (string line in raw.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            if (line.Length == 0)
            {
                // A blank line ends the header block; anything after it is body, not headers.
                Flush();
                break;
            }

            if (line[0] == ' ' || line[0] == '\t')
            {
                if (currentName != null)
                {
                    currentValue.Append(' ').Append(line.Trim());
                }

                continue;
            }

            int colon = line.IndexOf(':');

            if (colon <= 0)
            {
                // Not a header line and not a continuation. Skipped rather than invented into a
                // nameless entry.
                continue;
            }

            Flush();
            currentName = line[..colon].Trim();
            currentValue.Append(line[(colon + 1)..].Trim());
        }

        Flush();
        return headers;
    }

    // ── Meaning and notes ───────────────────────────────────────────────────

    private static string? DescribeMeaning(string name, object? value) => name switch
    {
        "lastVerbExecuted" when value is int verb => verb switch
        {
            0x100 => "read",
            0x102 => "submitted",
            0x104 => "replied to the sender",
            0x105 => "replied to all",
            0x106 => "forwarded",
            0x11C => "recalled",
            0x405 => "meeting accepted",
            0x406 => "meeting declined",
            0x407 => "meeting tentatively accepted",
            _ => $"verb 0x{verb:X}"
        },
        "messageFlags" when value is int flags => DescribeMessageFlags(flags),
        _ => null
    };

    private static string DescribeMessageFlags(int flags)
    {
        var names = new List<string>();

        if ((flags & 0x01) != 0) { names.Add("read"); }
        if ((flags & 0x02) != 0) { names.Add("unmodified"); }
        if ((flags & 0x04) != 0) { names.Add("submitted"); }
        if ((flags & 0x08) != 0) { names.Add("unsent"); }
        if ((flags & 0x10) != 0) { names.Add("has-attachment"); }
        if ((flags & 0x20) != 0) { names.Add("from-me"); }
        if ((flags & 0x40) != 0) { names.Add("associated"); }
        if ((flags & 0x80) != 0) { names.Add("resend"); }

        return names.Count == 0 ? "no flags set" : string.Join(", ", names);
    }

    /// <summary>
    /// Explains an absence that would otherwise look like a defect. Both cases here are documented
    /// MAPI behaviour rather than anything going wrong.
    /// </summary>
    private static string? DescribeAbsence(string name, string status, string? value)
    {
        if (name == "senderSmtpAddress" && string.Equals(value, "Unknown", StringComparison.Ordinal))
        {
            return "The MAPI spooler stores the literal string 'Unknown' when no transport provider "
                + "supplied a sender address. This is not a real address.";
        }

        if (status == "ok")
        {
            return null;
        }

        return name switch
        {
            "sentRepresentingSmtpAddress" when status is "not-present" or "empty" =>
                "Documented behaviour: the transport sets this on the outbound copy of a message and "
                + "leaves it unset on the local one, so it is usually absent on a sent item.",
            "inReplyToId" when status is "not-present" or "empty" =>
                "Set only on replies, so an original message does not carry it.",
            "internetMessageId" when status is "not-present" or "empty" =>
                "Set by the transport, so an item that never left the client does not carry it.",
            _ => null
        };
    }

    private static string DescribeStatus(string status, string? message) => status switch
    {
        "not-present" => "The item does not carry this property. That is an ordinary outcome for a "
            + "MAPI property, not a failure.",
        "empty" => "The item carries this property with an empty value. Exchange returns an empty "
            + "string rather than reporting the property missing for some tags on an item that never "
            + "traversed a transport, so this means the same thing as absent: there is no value here.",
        "blocked" => "Outlook refused to return this property. A security prompt that no program can "
            + "answer, or a permissions failure, withheld a value that exists.",
        "unsupported-or-blocked" => "Outlook returned MAPI_E_NOT_SUPPORTED. That means either the "
            + "property accessor cannot handle this property's type - PT_OBJECT properties cannot be "
            + "read at all - or Outlook's security prompt refused it. The two are indistinguishable "
            + "from the HRESULT alone.",
        _ => message ?? "The property could not be read."
    };

    // ── Item plumbing ───────────────────────────────────────────────────────

    /// <summary>
    /// Holds every COM object acquired while resolving an item so that all of them are released
    /// together, children before parents, whatever happens in between.
    /// </summary>
    private sealed class ItemScope
    {
        private Outlook.Inspector? _inspector;
        private Outlook.Explorer? _explorer;
        private Outlook.Selection? _selection;
        private object? _item;

        [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
        public object? Resolve(
            Outlook.Application application,
            Outlook.NameSpace session,
            string? entryId,
            string? storeId,
            bool useActiveItem)
        {
            if (!string.IsNullOrWhiteSpace(entryId))
            {
                _item = session.GetItemFromID(
                    entryId,
                    string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
                return _item;
            }

            if (!useActiveItem)
            {
                return null;
            }

            _inspector = application.ActiveInspector();

            if (_inspector != null)
            {
                _item = _inspector.CurrentItem;

                if (_item != null)
                {
                    return _item;
                }
            }

            _explorer = application.ActiveExplorer();

            if (_explorer != null)
            {
                _selection = _explorer.Selection;

                if (_selection != null && _selection.Count > 0)
                {
                    _item = _selection[1];
                    return _item;
                }
            }

            return null;
        }

        public void Release()
        {
            OutlookInteropRunner.ReleaseComObject(ref _item);
            OutlookInteropRunner.ReleaseComObject(ref _selection);
            OutlookInteropRunner.ReleaseComObject(ref _explorer);
            OutlookInteropRunner.ReleaseComObject(ref _inspector);
        }
    }

    /// <summary>
    /// <c>PropertyAccessor</c> exists on every Outlook item type but there is no common interface
    /// carrying it, so the type has to be tested. A switch is used rather than <c>dynamic</c> so
    /// that an unknown item class is reported honestly instead of throwing at the call site.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.PropertyAccessor? GetPropertyAccessor(object item) => item switch
    {
        Outlook.MailItem mail => mail.PropertyAccessor,
        Outlook.AppointmentItem appointment => appointment.PropertyAccessor,
        Outlook.MeetingItem meeting => meeting.PropertyAccessor,
        Outlook.ContactItem contact => contact.PropertyAccessor,
        Outlook.DistListItem distributionList => distributionList.PropertyAccessor,
        Outlook.TaskItem task => task.PropertyAccessor,
        Outlook.JournalItem journal => journal.PropertyAccessor,
        Outlook.NoteItem note => note.PropertyAccessor,
        Outlook.PostItem post => post.PropertyAccessor,
        Outlook.ReportItem report => report.PropertyAccessor,
        Outlook.SharingItem sharing => sharing.PropertyAccessor,
        _ => null
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.UserProperties? GetUserProperties(object item) => item switch
    {
        Outlook.MailItem mail => mail.UserProperties,
        Outlook.AppointmentItem appointment => appointment.UserProperties,
        Outlook.MeetingItem meeting => meeting.UserProperties,
        Outlook.ContactItem contact => contact.UserProperties,
        Outlook.DistListItem distributionList => distributionList.UserProperties,
        Outlook.TaskItem task => task.UserProperties,
        Outlook.JournalItem journal => journal.UserProperties,
        Outlook.NoteItem => null,
        Outlook.PostItem post => post.UserProperties,
        Outlook.ReportItem report => report.UserProperties,
        Outlook.SharingItem sharing => sharing.UserProperties,
        _ => null
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string DescribeItemType(object item) => item switch
    {
        Outlook.MailItem => "MailItem",
        Outlook.AppointmentItem => "AppointmentItem",
        Outlook.MeetingItem => "MeetingItem",
        Outlook.ContactItem => "ContactItem",
        Outlook.DistListItem => "DistListItem",
        Outlook.TaskItem => "TaskItem",
        Outlook.JournalItem => "JournalItem",
        Outlook.NoteItem => "NoteItem",
        Outlook.PostItem => "PostItem",
        Outlook.ReportItem => "ReportItem",
        Outlook.SharingItem => "SharingItem",
        _ => "unknown"
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? GetSubject(object item) => item switch
    {
        Outlook.MailItem mail => mail.Subject,
        Outlook.AppointmentItem appointment => appointment.Subject,
        Outlook.MeetingItem meeting => meeting.Subject,
        Outlook.ContactItem contact => contact.Subject,
        Outlook.TaskItem task => task.Subject,
        Outlook.PostItem post => post.Subject,
        Outlook.ReportItem report => report.Subject,
        Outlook.SharingItem sharing => sharing.Subject,
        _ => null
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? GetEntryId(object item) => item switch
    {
        Outlook.MailItem mail => mail.EntryID,
        Outlook.AppointmentItem appointment => appointment.EntryID,
        Outlook.MeetingItem meeting => meeting.EntryID,
        Outlook.ContactItem contact => contact.EntryID,
        Outlook.DistListItem distributionList => distributionList.EntryID,
        Outlook.TaskItem task => task.EntryID,
        Outlook.JournalItem journal => journal.EntryID,
        Outlook.NoteItem note => note.EntryID,
        Outlook.PostItem post => post.EntryID,
        Outlook.ReportItem report => report.EntryID,
        Outlook.SharingItem sharing => sharing.EntryID,
        _ => null
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static UserPropertyInfo DescribeUserProperty(Outlook.UserProperty property, int index)
    {
        var info = new UserPropertyInfo
        {
            Name = NullIfBlank(SafeGet(() => property.Name)) ?? $"(user property {index})",
            Type = DescribeUserPropertyType(property)
        };

        object? value;

        try
        {
            value = property.Value;
        }
        catch (COMException ex)
        {
            // A computed property - a formula or a combination field - can fail to evaluate. The
            // property still exists and is worth listing; only its value is missing.
            info.Note = OutlookInteropRunner.IsObjectModelGuardDenial(ex)
                ? "Outlook's security prompt refused this property's value."
                : $"This property's value could not be read: {ex.Message}";
            return info;
        }

        switch (value)
        {
            case null:
                break;

            case string text:
                info.Value = text;
                info.ValueType = "string";
                break;

            case bool flag:
                info.Value = flag ? "true" : "false";
                info.ValueType = "bool";
                break;

            case DateTime moment:
                info.Value = moment.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                info.ValueType = "dateTime";
                break;

            default:
                info.Value = Convert.ToString(value, CultureInfo.InvariantCulture);
                info.ValueType = value.GetType().Name;
                break;
        }

        return info;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? DescribeUserPropertyType(Outlook.UserProperty property)
    {
        try
        {
            return property.Type switch
            {
                Outlook.OlUserPropertyType.olOutlookInternal => "outlook-internal",
                Outlook.OlUserPropertyType.olText => "text",
                Outlook.OlUserPropertyType.olNumber => "number",
                Outlook.OlUserPropertyType.olDateTime => "date-time",
                Outlook.OlUserPropertyType.olYesNo => "yes-no",
                Outlook.OlUserPropertyType.olDuration => "duration",
                Outlook.OlUserPropertyType.olKeywords => "keywords",
                Outlook.OlUserPropertyType.olPercent => "percent",
                Outlook.OlUserPropertyType.olCurrency => "currency",
                Outlook.OlUserPropertyType.olFormula => "formula",
                Outlook.OlUserPropertyType.olCombination => "combination",
                Outlook.OlUserPropertyType.olInteger => "integer",
                Outlook.OlUserPropertyType.olEnumeration => "enumeration",
                Outlook.OlUserPropertyType.olSmartFrom => "smart-from",
                _ => "unknown"
            };
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static string BuildUnresolvedItemMessage(string? entryId)
        => string.IsNullOrWhiteSpace(entryId)
            ? "No Outlook item is currently open or selected, and no entryId was given."
            : "The requested Outlook item could not be resolved. An entryId is only meaningful "
              + "together with the storeId it came from.";

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
