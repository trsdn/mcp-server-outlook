using System.Text.Json;
using OutlookMcp.ComInterop.ServiceClient;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Listing rows projected from an Outlook <c>Table</c> rowset rather than by opening a message each
/// (#27).
///
/// <para>
/// The risk this change carries is not that it is slow - it is that a rowset silently reports less
/// than an opened item did, so a listing keeps saying <c>success: true</c> while quietly answering a
/// different question. Every test here is aimed at that: the central one compares the two
/// projections field for field over the same folder, so a column that reads as null, or comes back
/// in UTC, or renders a conversation id in the wrong case, fails rather than passing unnoticed.
/// </para>
///
/// <para>
/// Nothing belonging to the user is touched. Each test that needs a message of its own creates a
/// draft with a unique subject and deletes it in a <c>finally</c>.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailProjection")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailTableProjectionTests(ITestOutputHelper output)
{
    private const string Marker = "OutlookMcp-projection";

    /// <summary>
    /// The change itself: an ordinary listing must not open a single message. Asserted through the
    /// response rather than by timing, which would be flaky, and rather than not at all, which is how
    /// a performance change quietly stops being applied.
    /// </summary>
    [SkippableFact]
    public void MailList_ByDefault_IsAnsweredByTheTableProjection()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().List(folder: "drafts", maxCount: 5);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("table", result.Projection);
    }

    /// <summary>
    /// A rowset knows whether a message has attachments, never how many. The count must therefore be
    /// absent rather than reported as zero: a caller reading "0 attachments" off a message with three
    /// has been told something false under a successful response, which is worse than being told
    /// nothing.
    /// </summary>
    [SkippableFact]
    public void MailList_TableProjection_ReportsHasAttachmentAndOmitsTheExactCount()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string subject = UniqueSubject();

        try
        {
            draftId = CreateDraft(commands, subject);
            AttachAFile(draftId);

            MailSummaryInfo mine = FindListed(commands, subject, includeBodyPreview: false, out MailListResult listed);

            Assert.Equal("table", listed.Projection);
            Assert.True(mine.HasAttachment, "The draft has an attachment but the listing did not say so.");
            Assert.Null(mine.AttachmentCount);

            // On the wire too: a field that is null in the object but serialised as 0 would still
            // mislead the caller, which is the only place this actually matters.
            string json = JsonSerializer.Serialize(mine, ServiceProtocol.JsonOptions);
            output.WriteLine(json);
            Assert.DoesNotContain("\"attachmentCount\"", json, StringComparison.Ordinal);
            Assert.Contains("\"hasAttachment\":true", json, StringComparison.Ordinal);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The escape hatch has to actually work: asking for a body preview must open the messages, and
    /// once they are open the exact count is available again.
    /// </summary>
    [SkippableFact]
    public void MailList_WithIncludeBodyPreview_OpensItemsAndKeepsTheExactAttachmentCount()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string subject = UniqueSubject();

        try
        {
            draftId = CreateDraft(commands, subject);
            AttachAFile(draftId);

            MailSummaryInfo mine = FindListed(commands, subject, includeBodyPreview: true, out MailListResult listed);

            Assert.Equal("item", listed.Projection);
            Assert.True(mine.HasAttachment);
            Assert.Equal(1, mine.AttachmentCount);
            Assert.False(string.IsNullOrWhiteSpace(mine.BodyPreview));
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The regression that would matter most. A rowset cannot read a message body, so a free-text
    /// search answered from one would stop finding terms that appear only in the body - and would
    /// report that as "no such mail". The client-side scan mode must therefore still open items.
    /// </summary>
    [SkippableFact]
    public void MailSearch_WithAClientScanQuery_StillFindsATermThatAppearsOnlyInTheBody()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string subject = UniqueSubject();
        string bodyOnlyToken = "bodyonly" + Guid.NewGuid().ToString("N")[..12];

        try
        {
            var draft = commands.CreateMailDraft(subject: subject, body: "carrier " + bodyOnlyToken + " sentinel");
            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            var found = commands.Search(
                query: bodyOnlyToken,
                folder: "drafts",
                maxCount: 50,
                searchMode: "clientScan");

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Equal("item", found.Projection);
            Assert.Contains(found.Messages, m => m.EntryId == draftId);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The columns most likely to come back empty or wrong from a rowset, checked against values this
    /// test wrote itself so "null" cannot be mistaken for "correct". Categories in particular cannot
    /// be read through the DASL keywords property at all, and a conversation id has to be hex-encoded
    /// from a binary column - both would fail silently.
    /// </summary>
    [SkippableFact]
    public void MailList_TableProjection_ReportsCategoriesFlagDueDateAndConversationId()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string subject = UniqueSubject();

        try
        {
            draftId = CreateDraft(commands, subject);

            var categorised = commands.SetCategories(categories: "Blue Category", entryId: draftId, useActiveMail: false);
            Assert.True(categorised.Success, categorised.ErrorMessage);

            DateTime due = DateTime.Today.AddDays(5);
            var flagged = commands.SetFlag(
                entryId: draftId,
                useActiveMail: false,
                dueDate: due.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
                flagRequest: "Review");
            Assert.True(flagged.Success, flagged.ErrorMessage);

            MailSummaryInfo mine = FindListed(commands, subject, includeBodyPreview: false, out MailListResult listed);
            Assert.Equal("table", listed.Projection);

            Assert.Contains("Blue Category", mine.Categories);
            Assert.Equal("flagged", mine.FlagStatus);
            Assert.Equal("Review", mine.FlagRequest);
            Assert.NotNull(mine.FlagDueDate);
            Assert.Equal(due.Date, mine.FlagDueDate!.Value.Date);
            Assert.True(mine.IsDraft, "An unsent draft must still be reported as a draft.");
            Assert.Equal("mail", mine.ItemType);

            // Compared against what Outlook's own object model reports, not merely "not empty": a
            // lowercase or delimited rendering of the binary property would look plausible and would
            // never match a conversation.
            var read = commands.Read(entryId: draftId, useActiveMail: false);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.Equal(read.ConversationId, mine.ConversationId);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The whole contract in one test: the cheap projection and the expensive one must describe the
    /// same folder identically, message for message and field for field.
    ///
    /// <para>
    /// Blank and absent are treated as the same answer for the string fields. The two projections do
    /// spell an empty recipient list differently - a rowset omits it, an opened item returns an empty
    /// string - and that difference is real but immaterial; the claim under test is that the values
    /// agree, not that emptiness is spelled the same way.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void MailList_TableProjection_ReportsTheSameFieldsAsOpeningEachMessage()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();

        var projected = commands.List(folder: "drafts", maxCount: 15, includeBodyPreview: false);
        var hydrated = commands.List(folder: "drafts", maxCount: 15, includeBodyPreview: true);

        Assert.True(projected.Success, projected.ErrorMessage);
        Assert.True(hydrated.Success, hydrated.ErrorMessage);
        Assert.Equal("table", projected.Projection);
        Assert.Equal("item", hydrated.Projection);

        Skip.If(hydrated.ReturnedCount == 0, "Drafts is empty, so there is nothing to compare.");
        output.WriteLine($"comparing {projected.ReturnedCount} projected rows against {hydrated.ReturnedCount} opened items");

        Assert.Equal(
            hydrated.Messages.Select(m => m.EntryId),
            projected.Messages.Select(m => m.EntryId));

        foreach ((MailSummaryInfo fromTable, MailSummaryInfo fromItem) in projected.Messages.Zip(hydrated.Messages))
        {
            string where = $"entryId={fromItem.EntryId} subject='{fromItem.Subject}'";

            AssertSameText(fromItem.Subject, fromTable.Subject, "subject", where);
            AssertSameText(fromItem.SenderName, fromTable.SenderName, "senderName", where);
            AssertSameText(fromItem.SenderEmailAddress, fromTable.SenderEmailAddress, "senderEmailAddress", where);
            AssertSameText(fromItem.To, fromTable.To, "to", where);
            AssertSameText(fromItem.Cc, fromTable.Cc, "cc", where);
            AssertSameText(fromItem.ConversationId, fromTable.ConversationId, "conversationId", where);
            AssertSameText(fromItem.ConversationTopic, fromTable.ConversationTopic, "conversationTopic", where);
            AssertSameText(fromItem.StoreId, fromTable.StoreId, "storeId", where);
            AssertSameText(fromItem.FlagRequest, fromTable.FlagRequest, "flagRequest", where);

            Assert.Equal(fromItem.Categories, fromTable.Categories);
            Assert.Equal(fromItem.Unread, fromTable.Unread);
            Assert.Equal(fromItem.IsDraft, fromTable.IsDraft);
            Assert.Equal(fromItem.Importance, fromTable.Importance);
            Assert.Equal(fromItem.FlagStatus, fromTable.FlagStatus);
            Assert.Equal(fromItem.FlagDueDate, fromTable.FlagDueDate);
            Assert.Equal(fromItem.ItemType, fromTable.ItemType);
            Assert.Equal(fromItem.AttachmentCount > 0, fromTable.HasAttachment);

            // The ordering the paging cursor is a keyset walk over. A rowset that returned this in
            // UTC would shift every timestamp by the machine's offset and corrupt paging silently.
            Assert.Equal(fromItem.ReceivedTime, fromTable.ReceivedTime);
            Assert.Equal(fromItem.SentOn, fromTable.SentOn);
        }
    }

    /// <summary>
    /// The counting fields have to keep meaning what they meant, because a caller decides whether to
    /// keep paging from them. In particular a truncated page must still say so.
    /// </summary>
    [SkippableFact]
    public void MailList_TableProjection_ReportsTruncationAndOffersAContinuation()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var page = commands.List(folder: "drafts", maxCount: 1);

        Assert.True(page.Success, page.ErrorMessage);
        Assert.Equal("table", page.Projection);
        Skip.If(page.TotalItemCount < 3, "Drafts holds too few items to be truncated by a one-item page.");

        Assert.True(page.Truncated, "A one-item page of a multi-item folder must report truncation.");
        Assert.True(page.HasMore);
        Assert.False(string.IsNullOrWhiteSpace(page.NextCursor));
        Assert.Equal(1, page.ReturnedCount);
    }

    private static void AssertSameText(string? expected, string? actual, string field, string where)
    {
        string? left = string.IsNullOrWhiteSpace(expected) ? null : expected;
        string? right = string.IsNullOrWhiteSpace(actual) ? null : actual;

        Assert.True(
            string.Equals(left, right, StringComparison.Ordinal),
            $"{field} differs between projections: opened item gave '{expected}', table gave '{actual}' ({where}).");
    }

    private MailSummaryInfo FindListed(
        MailCommands commands,
        string subject,
        bool includeBodyPreview,
        out MailListResult listed)
    {
        listed = commands.List(
            folder: "drafts",
            maxCount: 50,
            includeBodyPreview: includeBodyPreview,
            subjectContains: subject);

        Assert.True(listed.Success, listed.ErrorMessage);
        output.WriteLine($"projection={listed.Projection} returned={listed.ReturnedCount} scanned={listed.ScannedCount}");

        MailSummaryInfo? mine = listed.Messages.FirstOrDefault(
            m => string.Equals(m.Subject, subject, StringComparison.Ordinal));

        Assert.True(mine != null, $"The listing did not return the draft this test created ('{subject}').");
        return mine!;
    }

    private static string UniqueSubject() => $"{Marker} {Guid.NewGuid():N}";

    private static string CreateDraft(MailCommands commands, string subject)
    {
        var draft = commands.CreateMailDraft(subject: subject, body: "projection placeholder body");
        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));
        return draft.EntryId!;
    }

    /// <summary>
    /// Attaches a real file through raw COM, deliberately not through this project's own attachment
    /// command: the point is to establish the fact independently of the code under test.
    /// </summary>
    private static void AttachAFile(string? entryId)
    {
        Assert.False(string.IsNullOrWhiteSpace(entryId));

        string path = Path.Combine(Path.GetTempPath(), $"outlookmcp-projection-{Guid.NewGuid():N}.txt");
        File.WriteAllText(path, "attachment payload");

        Assert.True(
            OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application),
            "Outlook stopped being available mid-test.");

        OutlookInterop.NameSpace? session = null;
        OutlookInterop.MailItem? mail = null;
        OutlookInterop.Attachments? attachments = null;

        try
        {
            session = application!.GetNamespace("MAPI");
            mail = session.GetItemFromID(entryId) as OutlookInterop.MailItem;
            Assert.NotNull(mail);

            attachments = mail!.Attachments;
            _ = attachments.Add(path, OutlookInterop.OlAttachmentType.olByValue, 1, "payload.txt");
            mail.Save();
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref attachments);
            OutlookInteropRunner.ReleaseComObject(ref mail);
            OutlookInteropRunner.ReleaseComObject(ref session);
            OutlookInteropRunner.ReleaseSharedComObject(ref application);
            File.Delete(path);
        }
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
            output.WriteLine("Skipping Outlook projection test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
