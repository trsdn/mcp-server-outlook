using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Commands.MessageProperties;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Internet message headers, named MAPI properties and user properties (#15).
///
/// <para>
/// The trap these tests are written around: <b>a draft has no transport headers</b>. It never
/// traversed a transport, so <c>PR_TRANSPORT_MESSAGE_HEADERS</c> is simply absent on it. A header
/// test that creates its own draft and reads headers back would pass forever without ever proving
/// that a single header can be parsed. Every positive assertion here is made against a real
/// received message taken from the Inbox, and the test skips with a stated reason when the mailbox
/// holds none.
/// </para>
///
/// <para>
/// The draft is still used, deliberately, but for the opposite claim: that an absent property is a
/// <em>success</em> with <c>headersPresent: false</c>, not an error. Those two tests are a matched
/// pair, and neither is meaningful without the other.
/// </para>
///
/// <para>
/// Drafts created here carry a unique subject prefix and are deleted again, because this mailbox is
/// shared with other work.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MessageProperty")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMessagePropertyTests(ITestOutputHelper output)
{
    private const string SubjectPrefix = "OutlookMcpPropTest";

    /// <summary>PR_SUBJECT. Present on every item, which makes it the honest round-trip probe.</summary>
    private const string PrSubject = "http://schemas.microsoft.com/mapi/proptag/0x0037001F";

    /// <summary>PR_TRANSPORT_MESSAGE_HEADERS. Absent on anything that never left the client.</summary>
    private const string PrTransportMessageHeaders = "http://schemas.microsoft.com/mapi/proptag/0x007D001F";

    /// <summary>PR_SENDER_ENTRYID, a PT_BINARY property, used to exercise the binary path.</summary>
    private const string PrSenderEntryId = "http://schemas.microsoft.com/mapi/proptag/0x0C190102";

    // ── Headers, against a genuinely received message ───────────────────────

    /// <summary>
    /// The test the whole feature turns on. A message that arrived over SMTP carries its transport
    /// headers, and they must come back parsed into names and values rather than as one blob.
    /// </summary>
    [SkippableFact]
    public void GetHeaders_ReceivedMessage_ReturnsParsedTransportHeaders()
    {
        EnsureOutlookAvailable();

        var (entryId, storeId, subject) = RequireReceivedMessageWithHeaders();

        var result = new PropertyCommands().GetHeaders(entryId: entryId, storeId: storeId);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.True(result.HeadersPresent, $"'{subject}' reported no transport headers.");
        Assert.NotEmpty(result.Headers);
        Assert.Equal(result.Headers.Count, result.HeaderCount);

        foreach (var header in result.Headers)
        {
            Assert.False(string.IsNullOrWhiteSpace(header.Name), "A header arrived without a name.");
            Assert.DoesNotContain(':', header.Name);
        }

        // Every message that traversed a transport has at least one Received header. Its absence
        // would mean the value was read but not parsed.
        Assert.Contains(result.Headers, h => h.Name.Equals("Received", StringComparison.OrdinalIgnoreCase));

        output.WriteLine($"'{subject}': {result.HeaderCount} headers.");
        foreach (var header in result.Headers.Take(12))
        {
            output.WriteLine($"  {header.Name}: {Truncate(header.Value)}");
        }
    }

    /// <summary>
    /// Folded continuation lines must be joined back onto the header they belong to. A naive
    /// line-by-line split produces phantom headers with no name and truncates long ones such as
    /// <c>Authentication-Results</c>, which is precisely the header an agent checking SPF or DKIM
    /// needs whole.
    /// </summary>
    [SkippableFact]
    public void GetHeaders_UnfoldsContinuationLines()
    {
        EnsureOutlookAvailable();

        var (entryId, storeId, _) = RequireReceivedMessageWithHeaders();

        var result = new PropertyCommands().GetHeaders(entryId: entryId, storeId: storeId, includeRaw: true);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.Raw);

        int foldedLines = result.Raw!
            .Replace("\r\n", "\n")
            .Split('\n')
            .Count(line => line.Length > 0 && (line[0] == ' ' || line[0] == '\t'));

        Skip.If(foldedLines == 0, "No header on this message is folded, so there is nothing to unfold.");

        // A folded line starts with whitespace and carries no colon-delimited name of its own. If
        // folding were ignored, those lines would surface as headers with an empty or absurd name.
        Assert.All(result.Headers, h => Assert.False(string.IsNullOrWhiteSpace(h.Name)));
        Assert.True(
            result.HeaderCount < result.Raw.Replace("\r\n", "\n").Split('\n').Length,
            "The header count equals the raw line count, so continuation lines were not folded in.");

        output.WriteLine($"{foldedLines} folded line(s) collapsed into {result.HeaderCount} headers.");
    }

    /// <summary>
    /// Header blocks run to tens of kilobytes. Asking for one header by name has to return that
    /// header rather than everything.
    /// </summary>
    [SkippableFact]
    public void GetHeaders_WithHeaderName_ReturnsOnlyThatHeader()
    {
        EnsureOutlookAvailable();

        var (entryId, storeId, _) = RequireReceivedMessageWithHeaders();

        var result = new PropertyCommands().GetHeaders(entryId: entryId, storeId: storeId, headerName: "Received");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.HeadersPresent);
        Assert.NotEmpty(result.Headers);
        Assert.All(result.Headers, h => Assert.Equal("Received", h.Name, ignoreCase: true));

        output.WriteLine($"{result.HeaderCount} Received header(s) returned.");
    }

    /// <summary>
    /// The matched pair to the test above. A draft never traversed a transport, so it has no
    /// headers - and that is an answer, not a failure. Reporting it as an error would make "this
    /// message was composed locally" indistinguishable from "Outlook refused the read".
    /// </summary>
    [SkippableFact]
    public void GetHeaders_Draft_ReportsAbsentHeadersAsSuccess()
    {
        EnsureOutlookAvailable();

        var mail = new MailCommands();
        string subject = UniqueSubject();
        var draft = mail.CreateMailDraft(subject: subject, body: "Header probe.");
        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.NotNull(draft.EntryId);

        try
        {
            var result = new PropertyCommands().GetHeaders(entryId: draft.EntryId, storeId: draft.StoreId);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.False(result.HeadersPresent, "A locally created draft reported transport headers.");
            Assert.Empty(result.Headers);
            Assert.False(string.IsNullOrWhiteSpace(result.Note), "No explanation was given for the absent headers.");

            output.WriteLine($"Draft note: {result.Note}");
        }
        finally
        {
            mail.Delete(entryId: draft.EntryId, storeId: draft.StoreId, useActiveMail: false);
        }
    }

    // ── The curated property set ────────────────────────────────────────────

    /// <summary>
    /// Every curated property must come back with a status, present or not. A property silently
    /// missing from the list would be indistinguishable from one that was never asked for.
    /// </summary>
    [SkippableFact]
    public void GetKnown_ReceivedMessage_ReportsEveryPropertyWithAStatus()
    {
        EnsureOutlookAvailable();

        var (entryId, storeId, subject) = RequireReceivedMessageWithHeaders();

        var result = new PropertyCommands().GetKnown(entryId: entryId, storeId: storeId);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.NotEmpty(result.Properties);

        foreach (var property in result.Properties)
        {
            Assert.False(string.IsNullOrWhiteSpace(property.Name), "A property arrived without a name.");
            Assert.False(string.IsNullOrWhiteSpace(property.Dasl), $"'{property.Name}' arrived without its DASL name.");
            Assert.False(string.IsNullOrWhiteSpace(property.Status), $"'{property.Name}' arrived without a status.");

            if (property.Found)
            {
                Assert.Equal("ok", property.Status);
                Assert.NotNull(property.Value);
            }
            else
            {
                Assert.NotEqual("ok", property.Status);
            }

            output.WriteLine($"{property.Name} [{property.Status}] = {Truncate(property.Value)}");
        }

        // A message that arrived over SMTP always has an Internet message id.
        var messageId = result.Properties.Single(p => p.Name == "internetMessageId");
        Assert.True(messageId.Found, $"'{subject}' arrived over SMTP but reports no Internet message id.");
        Assert.Contains('@', messageId.Value!);
    }

    // ── Arbitrary DASL reads ────────────────────────────────────────────────

    /// <summary>
    /// The round trip that proves arbitrary DASL reads actually work: PR_SUBJECT read through the
    /// property accessor must equal the subject the mail surface reports for the same item.
    /// </summary>
    [SkippableFact]
    public void GetProperty_WithSubjectTag_MatchesTheSubjectTheMailSurfaceReports()
    {
        EnsureOutlookAvailable();

        var (entryId, storeId, subject) = RequireAnyInboxMessage();

        var result = new PropertyCommands().GetProperty(PrSubject, entryId: entryId, storeId: storeId);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.True(result.Found, $"PR_SUBJECT was not readable on '{subject}'.");
        Assert.Equal("ok", result.Status);
        Assert.Equal("string", result.ValueType);
        Assert.Equal(subject, result.Value);
    }

    /// <summary>
    /// A binary property must not be stringified into whatever <c>ToString()</c> happens to give.
    /// It comes back as a byte array, and it is reported as one.
    /// </summary>
    [SkippableFact]
    public void GetProperty_WithBinaryTag_ReportsBytesRatherThanAToStringOfThem()
    {
        EnsureOutlookAvailable();

        var (entryId, storeId, subject) = RequireAnyInboxMessage();

        var result = new PropertyCommands().GetProperty(PrSenderEntryId, entryId: entryId, storeId: storeId);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(!result.Found, $"'{subject}' carries no sender entry id: {result.Status}.");

        Assert.Equal("binary", result.ValueType);
        Assert.NotNull(result.BinaryLength);
        Assert.True(result.BinaryLength > 0);
        Assert.False(string.IsNullOrWhiteSpace(result.Base64), "A binary property came back with no base64 value.");
        Assert.DoesNotContain("System.Byte[]", result.Value ?? string.Empty, StringComparison.Ordinal);

        output.WriteLine($"PR_SENDER_ENTRYID: {result.BinaryLength} bytes, hex={Truncate(result.Hex)}");
    }

    /// <summary>
    /// A property with no usable value is a success reporting <c>empty</c> or <c>not-present</c>,
    /// never an error.
    ///
    /// <para>
    /// Which of the two comes back was measured, not assumed. Asked for
    /// <c>PR_TRANSPORT_MESSAGE_HEADERS</c> on a draft, this Exchange store returns an <b>empty
    /// string</b> rather than raising <c>MAPI_E_NOT_FOUND</c>. Passing that through as a found
    /// value would answer "yes, this message has transport headers" and hand back nothing, so an
    /// empty value is never reported as <c>ok</c>.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void GetProperty_WithValuelessTag_SucceedsAndReportsNoUsableValue()
    {
        EnsureOutlookAvailable();

        var mail = new MailCommands();
        string subject = UniqueSubject();
        var draft = mail.CreateMailDraft(subject: subject, body: "Absent property probe.");
        Assert.True(draft.Success, draft.ErrorMessage);

        try
        {
            var result = new PropertyCommands().GetProperty(
                PrTransportMessageHeaders, entryId: draft.EntryId, storeId: draft.StoreId);

            output.WriteLine($"headers on a draft: status={result.Status}, note={result.Note}");

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.False(result.Found, "A draft reported a usable transport-headers value.");
            Assert.Equal("empty", result.Status);
            Assert.Null(result.Value);
            Assert.False(string.IsNullOrWhiteSpace(result.Note));
        }
        finally
        {
            mail.Delete(entryId: draft.EntryId, storeId: draft.StoreId, useActiveMail: false);
        }
    }

    /// <summary>
    /// A property tag that is well formed but that no item carries must come back as
    /// <c>not-present</c> on a successful call. Outlook raises <c>MAPI_E_NOT_FOUND</c> for this,
    /// and passing that through as a failure would make an ordinary absence look like a broken
    /// call.
    /// </summary>
    [SkippableFact]
    public void GetProperty_WithAbsentTag_SucceedsAndReportsNotPresent()
    {
        EnsureOutlookAvailable();

        var mail = new MailCommands();
        string subject = UniqueSubject();
        var draft = mail.CreateMailDraft(subject: subject, body: "Absent property probe.");
        Assert.True(draft.Success, draft.ErrorMessage);

        try
        {
            // PR_ATTACH_CONTENT_ID is an attachment property. It is well formed, it is a string,
            // and a message never carries it - which is exactly the "genuinely missing" case.
            var result = new PropertyCommands().GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                entryId: draft.EntryId,
                storeId: draft.StoreId);

            output.WriteLine($"attachment content id on a message: status={result.Status}");

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.False(result.Found);
            Assert.Equal("not-present", result.Status);
            Assert.Null(result.Value);
            Assert.False(string.IsNullOrWhiteSpace(result.Note));
        }
        finally
        {
            mail.Delete(entryId: draft.EntryId, storeId: draft.StoreId, useActiveMail: false);
        }
    }

    /// <summary>
    /// A string that is not a DASL property name at all is refused up front, with the accepted
    /// namespaces named. Handing it to Outlook produces "The property name is invalid", which tells
    /// the caller nothing about what a valid one looks like.
    /// </summary>
    [SkippableFact]
    public void GetProperty_WithMalformedDasl_IsRefusedWithGuidance()
    {
        EnsureOutlookAvailable();

        var result = new PropertyCommands().GetProperty("PR_SUBJECT");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Contains("schemas.microsoft.com", result.ErrorMessage);
    }

    [SkippableFact]
    public void GetProperty_WithNoDasl_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new PropertyCommands().GetProperty("   ");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
    }

    // ── User properties ─────────────────────────────────────────────────────

    /// <summary>
    /// An item with no custom properties reports an empty list as a success. Most items have none,
    /// so this is the common case and must not look like a failure.
    /// </summary>
    [SkippableFact]
    public void ListUserProperties_ItemWithNone_SucceedsWithAnEmptyList()
    {
        EnsureOutlookAvailable();

        var mail = new MailCommands();
        string subject = UniqueSubject();
        var draft = mail.CreateMailDraft(subject: subject, body: "User property probe.");
        Assert.True(draft.Success, draft.ErrorMessage);

        try
        {
            var result = new PropertyCommands().ListUserProperties(
                entryId: draft.EntryId, storeId: draft.StoreId);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(result.UserProperties.Count, result.Count);
        }
        finally
        {
            mail.Delete(entryId: draft.EntryId, storeId: draft.StoreId, useActiveMail: false);
        }
    }

    /// <summary>
    /// An entry id nobody has must be refused rather than quietly answering about the item the user
    /// happens to have selected in Outlook.
    /// </summary>
    [SkippableFact]
    public void GetHeaders_WithUnresolvableItem_IsRefusedRatherThanFallingBackToTheSelection()
    {
        EnsureOutlookAvailable();

        var result = new PropertyCommands().GetHeaders(
            entryId: "0000000000000000000000000000000000000000", useActiveMail: false);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
    }

    // ── Helpers ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Finds a message in the Inbox that genuinely carries transport headers. Scans rather than
    /// taking the first item, because an internally delivered item, a calendar response or a
    /// locally filed message may have none, and testing headers against one of those would prove
    /// nothing at all.
    /// </summary>
    private (string EntryId, string? StoreId, string Subject) RequireReceivedMessageWithHeaders()
    {
        var listing = new MailCommands().List(folder: "inbox", maxCount: 25);
        Skip.If(!listing.Success, listing.ErrorMessage ?? "The inbox could not be listed.");
        Skip.If(listing.Messages.Count == 0, "The inbox is empty, so there is no received message to read headers from.");

        var properties = new PropertyCommands();

        foreach (var item in listing.Messages)
        {
            if (string.IsNullOrWhiteSpace(item.EntryId))
            {
                continue;
            }

            var probe = properties.GetHeaders(entryId: item.EntryId, storeId: item.StoreId);

            if (probe is { Success: true, HeadersPresent: true } && probe.Headers.Count > 0)
            {
                output.WriteLine($"Using received message: '{item.Subject}' ({probe.HeaderCount} headers).");
                return (item.EntryId, item.StoreId, item.Subject ?? "(no subject)");
            }
        }

        Skip.If(true,
            "No message in the first 25 inbox items carries transport headers, so nothing here "
            + "traversed an SMTP transport. Header behaviour cannot be verified on this mailbox.");
        throw new InvalidOperationException("unreachable");
    }

    private static (string EntryId, string? StoreId, string Subject) RequireAnyInboxMessage()
    {
        var listing = new MailCommands().List(folder: "inbox", maxCount: 5);
        Skip.If(!listing.Success, listing.ErrorMessage ?? "The inbox could not be listed.");

        var item = listing.Messages.FirstOrDefault(i => !string.IsNullOrWhiteSpace(i.EntryId)
            && !string.IsNullOrWhiteSpace(i.Subject));

        Skip.If(item is null, "The inbox holds no message with both an entry id and a subject.");
        return (item!.EntryId!, item.StoreId, item.Subject!);
    }

    private static string UniqueSubject() => $"{SubjectPrefix}-{Guid.NewGuid():N}";

    private static string Truncate(string? value)
        => value is null ? "(none)" : value.Length <= 70 ? value : value[..70] + "...";

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
