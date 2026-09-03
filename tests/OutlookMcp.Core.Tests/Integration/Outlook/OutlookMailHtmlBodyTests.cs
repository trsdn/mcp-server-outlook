using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// HTML message bodies (#15), and the reply-flattening bug that writing them exposed.
///
/// <para>
/// Every assertion here reads the stored message back through <b>raw COM</b> rather than through this
/// project's own reader. That is deliberate. Round-tripping through our own read path would pass
/// happily if the write and the read shared a mistake - and "wrote plain text, read plain text back,
/// declared success" is precisely the failure this repository keeps finding. <c>HTMLBody</c> and
/// <c>BodyFormat</c> straight off the <c>MailItem</c> are the independent check.
/// </para>
///
/// <para>
/// Nothing is sent. Every message these tests write is a draft they created and delete again in a
/// <c>finally</c>. The one message they do not create is the received original used to test replying,
/// which is the user's real mail and is only ever read.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailHtmlBody")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailHtmlBodyTests(ITestOutputHelper output)
{
    private const string Marker = "OutlookMcp html-body test";

    /// <summary>
    /// The feature itself: markup asked for as HTML has to arrive as markup, not as text that looks
    /// like markup.
    /// </summary>
    [SkippableFact]
    public void SetBody_AsHtml_StoresRealMarkup()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            var set = commands.SetBody(
                body: "<p>Hello <b>world</b>.</p>",
                entryId: draftId,
                useActiveMail: false,
                bodyFormat: "html");

            Assert.True(set.Success, set.ErrorMessage);

            (string? html, string? plain, int format) = ReadRaw(draftId);

            // Markup, not text that looks like markup.
            Assert.Contains("<b>world</b>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("&lt;b&gt;", html, StringComparison.OrdinalIgnoreCase);

            // olFormatHTML. If this is still plain, Outlook took the string as text.
            Assert.Equal(2, format);

            // And the tags must not survive into the plain-text projection.
            Assert.DoesNotContain("<b>", plain, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("world", plain, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The other half, and the more dangerous one: text asked for as plain must never be interpreted.
    /// An agent relaying a user's words has no idea whether they contain angle brackets.
    /// </summary>
    [SkippableFact]
    public void SetBody_AsPlain_DoesNotInterpretMarkup()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            // Make it an HTML message first, so the plain write has something to overwrite. This is
            // the case where injection would actually be possible.
            var toHtml = commands.SetBody(
                body: "<p>original</p>", entryId: draftId, useActiveMail: false, bodyFormat: "html");
            Assert.True(toHtml.Success, toHtml.ErrorMessage);

            const string Literal = "compare <b>a</b> with <script>alert(1)</script>";

            var set = commands.SetBody(
                body: Literal, entryId: draftId, useActiveMail: false, bodyFormat: "plain");

            Assert.True(set.Success, set.ErrorMessage);

            (string? html, string? plain, _) = ReadRaw(draftId);

            // The user's words, intact and uninterpreted.
            Assert.Contains(Literal, plain, StringComparison.Ordinal);

            // Nothing live reached the HTML projection.
            Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// create-draft has to offer the same choice as set-body, or composing formatted mail means
    /// creating a draft and then immediately rewriting it.
    /// </summary>
    [SkippableFact]
    public void CreateDraft_AsHtml_StoresRealMarkup()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            var draft = commands.CreateMailDraft(
                subject: Marker,
                body: "<ul><li>first</li><li>second</li></ul>",
                bodyFormat: "html");

            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            (string? html, _, int format) = ReadRaw(draftId);

            Assert.Contains("<li>first</li>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(2, format);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// Backwards compatibility. Every existing caller omits the new argument, and must keep getting
    /// exactly what it got before: text, uninterpreted.
    /// </summary>
    [SkippableFact]
    public void CreateDraft_WithNoFormatArgument_IsStillPlainText()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            const string Literal = "not <b>markup</b>";

            var draft = commands.CreateMailDraft(subject: Marker, body: Literal);

            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            (_, string? plain, _) = ReadRaw(draftId);

            Assert.Contains(Literal, plain, StringComparison.Ordinal);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// A format this surface does not support has to be refused, not guessed at. Silently treating
    /// "richtext" as plain would put markup the caller expected to be rendered in front of a human as
    /// visible tag soup.
    /// </summary>
    [SkippableFact]
    public void SetBody_WithAnUnsupportedFormat_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            var set = commands.SetBody(
                body: "anything", entryId: draftId, useActiveMail: false, bodyFormat: "richtext");

            Assert.False(set.Success);
            Assert.Contains("richtext", set.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The bug this slice found. Replying used to read the draft's plain-text <c>Body</c> - which is
    /// a lossy projection of the quoted original - and write it straight back, so the whole quoted
    /// thread arrived at the recipient as flattened plain text. It reported success, and the text was
    /// all still there, so nothing looked wrong until you opened the draft.
    /// </summary>
    [SkippableFact]
    public void Reply_ToAnHtmlMessage_KeepsTheQuotedThreadAsHtml()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? replyId = null;

        try
        {
            string originalId = FindReceivedHtmlMessage(commands);

            (string? originalHtml, _, _) = ReadRaw(originalId);

            // Whatever rich markup the original carried, some of it has to survive into the quote.
            //
            // Asserting only "the reply is still BodyFormat=HTML" is not enough, and that was the
            // first version of this test: writing plain text into an HTML draft leaves the format
            // alone and makes Outlook regenerate a wrapper around the flattened text, so the
            // assertion passed against the very bug it was written to catch. Found by reinstating
            // the bug deliberately and watching this test stay green.
            var richness = RichMarkupIn(originalHtml);

            Skip.If(richness.Count == 0, "The chosen original carries no rich markup to lose.");

            var reply = commands.Reply(
                entryId: originalId,
                useActiveMail: false,
                body: "Thanks - noted.");

            Assert.True(reply.Success, reply.ErrorMessage);
            replyId = reply.EntryId;

            (string? html, string? plain, int format) = ReadRaw(replyId);

            // Still an HTML message.
            Assert.Equal(2, format);

            var survived = RichMarkupIn(html);

            output.WriteLine($"original markup: {string.Join(", ", richness)}");
            output.WriteLine($"survived in reply: {string.Join(", ", survived)}");

            Assert.True(
                survived.Overlaps(richness),
                $"None of the original's markup ({string.Join(", ", richness)}) survived into the "
                + "reply, so the quoted thread was flattened to plain text.");

            // And the caller's own words are there, above the quote.
            Assert.Contains("Thanks - noted.", plain, StringComparison.Ordinal);
        }
        finally
        {
            DeleteQuietly(commands, replyId);
        }
    }

    /// <summary>
    /// Prepending plain text into an HTML reply means escaping it. If it went in raw, a user writing
    /// "profit &lt; loss" would silently lose everything after the bracket.
    /// </summary>
    [SkippableFact]
    public void Reply_WithPlainTextContainingBrackets_EscapesItIntoTheHtml()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? replyId = null;

        try
        {
            string originalId = FindReceivedHtmlMessage(commands);

            const string Literal = "profit < loss & <b>margin</b>";

            var reply = commands.Reply(entryId: originalId, useActiveMail: false, body: Literal);

            Assert.True(reply.Success, reply.ErrorMessage);
            replyId = reply.EntryId;

            (string? html, string? plain, _) = ReadRaw(replyId);

            // Survives to the reader intact...
            Assert.Contains(Literal, plain, StringComparison.Ordinal);

            // ...because it was escaped rather than injected.
            Assert.Contains("&lt;b&gt;margin&lt;/b&gt;", html, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteQuietly(commands, replyId);
        }
    }

    /// <summary>
    /// The point of allowing HTML on a reply: the caller's markup renders, and the quoted original
    /// underneath it still survives.
    /// </summary>
    [SkippableFact]
    public void Reply_AsHtml_PutsRenderedMarkupAboveTheQuote()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? replyId = null;

        try
        {
            string originalId = FindReceivedHtmlMessage(commands);

            var reply = commands.Reply(
                entryId: originalId,
                useActiveMail: false,
                body: "<p>See <b>attached figures</b>.</p>",
                bodyFormat: "html");

            Assert.True(reply.Success, reply.ErrorMessage);
            replyId = reply.EntryId;

            (string? html, string? plain, _) = ReadRaw(replyId);

            Assert.Contains("<b>attached figures</b>", html, StringComparison.OrdinalIgnoreCase);

            // The quote is still below it.
            Assert.True(
                plain!.Length > "See attached figures.".Length + 20,
                "The quoted original appears to have been replaced rather than kept.");
        }
        finally
        {
            DeleteQuietly(commands, replyId);
        }
    }

    /// <summary>
    /// Forward takes the same argument as reply. Worth its own test because forward builds its draft
    /// through a different Outlook call, and the two have differed before.
    /// </summary>
    [SkippableFact]
    public void Forward_AsHtml_KeepsTheForwardedThreadAsHtml()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? forwardId = null;

        try
        {
            string originalId = FindReceivedHtmlMessage(commands);

            var forward = commands.Forward(
                entryId: originalId,
                useActiveMail: false,
                body: "<p>Forwarding <i>for information</i>.</p>",
                bodyFormat: "html");

            Assert.True(forward.Success, forward.ErrorMessage);
            forwardId = forward.EntryId;

            (string? html, _, int format) = ReadRaw(forwardId);

            Assert.Equal(2, format);
            Assert.Contains("<i>for information</i>", html, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            DeleteQuietly(commands, forwardId);
        }
    }

    /// <summary>
    /// The kinds of markup that carry meaning a reader would miss if they vanished, <b>and that
    /// Outlook cannot invent</b>.
    ///
    /// <para>
    /// Deliberately excludes <c>&lt;p&gt;</c>, <c>&lt;br&gt;</c>, <c>&lt;div&gt;</c>,
    /// <c>&lt;span&gt;</c> and <c>&lt;style&gt;</c>, which Outlook emits when it wraps flattened
    /// plain text. Also excludes <c>&lt;a href&gt;</c>, which looks like the strongest signal of all
    /// and is in fact useless: <b>Outlook auto-linkifies bare URLs</b> when it converts plain text to
    /// HTML, so links reappear in a body that has just been stripped of everything else. That was
    /// measured, not assumed - with the flattening bug reinstated, an original carrying
    /// <c>&lt;a&gt;</c>, <c>&lt;img&gt;</c>, <c>&lt;b&gt;</c> and <c>&lt;strong&gt;</c> came back
    /// with the images and the emphasis gone and the links intact, which was enough to keep an
    /// earlier version of this assertion green.
    /// </para>
    /// </summary>
    private static readonly string[] RichTags =
        ["<table", "<img", "<ul", "<ol", "<li", "<b>", "<strong", "<i>", "<em", "<h1", "<h2", "<h3"];

    private static HashSet<string> RichMarkupIn(string? html)
    {
        var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrEmpty(html))
        {
            return found;
        }

        foreach (string tag in RichTags)
        {
            if (html.Contains(tag, StringComparison.OrdinalIgnoreCase))
            {
                found.Add(tag);
            }
        }

        return found;
    }

    /// <summary>
    /// Reading the message back through raw COM, deliberately bypassing this project's own reader.
    /// Returns the HTML body, the plain-text projection, and <c>BodyFormat</c> as an integer
    /// (1 = plain, 2 = HTML, 3 = RTF).
    /// </summary>
    private static (string? Html, string? Plain, int Format) ReadRaw(string? entryId)
    {
        Assert.False(string.IsNullOrWhiteSpace(entryId), "No entry id to read back.");

        Assert.True(
            OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application),
            "Outlook stopped being available mid-test.");

        OutlookInterop.NameSpace? session = null;
        OutlookInterop.MailItem? mail = null;

        try
        {
            session = application!.GetNamespace("MAPI");
            mail = session.GetItemFromID(entryId) as OutlookInterop.MailItem;

            Assert.NotNull(mail);

            return (mail!.HTMLBody, mail.Body, (int)mail.BodyFormat);
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref mail);
            OutlookInteropRunner.ReleaseComObject(ref session);
            OutlookInteropRunner.ReleaseComObject(ref application);
        }
    }

    /// <summary>
    /// Finds a received message that is actually an HTML one. Skips rather than fails if the mailbox
    /// has none: a plain-text-only mailbox cannot demonstrate the flattening bug either way, and
    /// pretending otherwise would make the test pass without having tested anything.
    /// </summary>
    private string FindReceivedHtmlMessage(MailCommands commands)
    {
        var listed = commands.List(folder: "inbox", maxCount: 25);

        if (listed.Success)
        {
            foreach (var candidate in listed.Messages.Where(
                m => !m.IsDraft
                     && !string.IsNullOrWhiteSpace(m.EntryId)
                     && (m.ItemType == null || m.ItemType == "mail")))
            {
                (_, _, int format) = ReadRaw(candidate.EntryId);

                if (format == 2)
                {
                    output.WriteLine($"Replying to received message '{candidate.Subject}'.");
                    return candidate.EntryId!;
                }
            }
        }

        throw new SkipException("This mailbox holds no received HTML message to reply to.");
    }

    private static string CreateDraft(MailCommands commands)
    {
        var draft = commands.CreateMailDraft(subject: Marker, body: "placeholder");
        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));
        return draft.EntryId!;
    }

    private static void DeleteQuietly(MailCommands commands, string? entryId)
    {
        if (!string.IsNullOrWhiteSpace(entryId))
        {
            _ = commands.Delete(entryId: entryId, useActiveMail: false);
        }
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping Outlook html-body test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseComObject(ref application);
    }
}
