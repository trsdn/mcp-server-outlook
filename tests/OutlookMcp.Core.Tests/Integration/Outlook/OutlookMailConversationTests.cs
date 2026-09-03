using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Conversation / thread support against a real mailbox (#39).
///
/// <para>
/// The acceptance criterion these exist for is the one an agent actually depends on: <b>a reply and
/// the message it replies to must come back from a single call, as one thread</b>. Asserting only
/// that "a conversation id was returned" would pass against an implementation that returned a
/// one-item thread for every message, which is exactly the useless answer #39 exists to remove.
/// </para>
///
/// <para>
/// The thread is established with a real received message plus a reply draft to it. Outlook refuses
/// to build a reply from an unsent draft (#92), so a draft-to-draft pair - the obvious design, and
/// the one first written here - cannot work at all. A reply draft still shares its parent's
/// conversation, so the relationship is real and <b>nothing is sent</b> to establish it. Only the
/// reply draft is deleted afterwards; the received original is the user's own mail.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailConversation")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailConversationTests(ITestOutputHelper output)
{
    /// <summary>
    /// The whole point of the feature: one call, both messages.
    /// </summary>
    [SkippableFact]
    public void GetConversation_ForAMessageAndItsReply_ReturnsBothInOneThread()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? replyId = null;

        try
        {
            // The original has to be a message that was actually received. Outlook cannot build a
            // reply from an unsent draft (#92), so the earlier draft-to-draft version of this test
            // could never have passed - it had simply never run, because #90 made every Outlook test
            // skip itself. A reply draft still shares its parent's conversation, so the thread
            // relationship is real and nothing is sent to establish it.
            string originalId = FindReceivedMessage(commands);
            replyId = CreateReply(commands, originalId);

            var thread = ReadThreadContaining(commands, originalId, replyId);

            Assert.True(thread.Success, thread.ErrorMessage);
            output.WriteLine($"thread '{thread.ConversationTopic}' returned {thread.ReturnedCount} item(s).");

            var ids = thread.Messages.Select(m => m.EntryId).ToList();
            Assert.Contains(originalId, ids);
            Assert.Contains(replyId, ids);
        }
        finally
        {
            // Only the reply is ours to delete. The original is the user's real mail.
            DeleteQuietly(commands, replyId);
        }
    }

    /// <summary>
    /// A thread is only readable if it is in reading order. Asserted explicitly because "returns the
    /// right items" and "returns them in a usable order" are different claims and the first can pass
    /// while the second is false.
    /// </summary>
    [SkippableFact]
    public void GetConversation_ReturnsItemsOldestFirst()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? replyId = null;

        try
        {
            string originalId = FindReceivedMessage(commands);
            replyId = CreateReply(commands, originalId);

            var thread = ReadThreadContaining(commands, originalId, replyId);
            Assert.True(thread.Success, thread.ErrorMessage);

            Assert.Equal("receivedTime", thread.SortedBy);
            Assert.Equal("ascending", thread.SortDirection);

            var times = thread.Messages
                .Select(m => m.ReceivedTime ?? m.SentOn)
                .Where(t => t.HasValue)
                .Select(t => t!.Value)
                .ToList();

            Assert.Equal(times.OrderBy(t => t).ToList(), times);
        }
        finally
        {
            DeleteQuietly(commands, replyId);
        }
    }

    /// <summary>
    /// A thread spans folders - a reply lives in Drafts or Sent Items while the original sits in the
    /// Inbox - so each item has to say where it is. Without this a caller cannot act on a thread item
    /// at all beyond reading it.
    /// </summary>
    [SkippableFact]
    public void GetConversation_ItemsReportTheFolderTheyLiveIn()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? originalId = null;

        try
        {
            originalId = CreateDraft(commands, $"OutlookMcp folder {Guid.NewGuid():N}");

            var thread = commands.GetConversation(entryId: originalId, useActiveMail: false);
            Assert.True(thread.Success, thread.ErrorMessage);
            Assert.NotEmpty(thread.Messages);

            Assert.All(thread.Messages, m => Assert.False(string.IsNullOrWhiteSpace(m.FolderPath)));
        }
        finally
        {
            DeleteQuietly(commands, originalId);
        }
    }

    /// <summary>
    /// A conversation is not all mail. On the mailbox this was written against, an ordinary thread
    /// held seven items of which only three were messages: the rest were a meeting invitation, the
    /// calendar appointment it created, and the acceptance. The invitation and the acceptance are
    /// frequently the substance of the thread - when did we agree to meet, and did they say yes.
    ///
    /// <para>
    /// Those items used to be reduced to a number, so an agent asked to summarise that conversation
    /// saw three replies and the digit 4. They must come back identified.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void GetConversation_ReturnsNonMailThreadItemsRatherThanOnlyCountingThem()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var thread = FindThreadWithNonMailItems(commands);

        Skip.If(thread is null, "No conversation in this inbox contains a non-mail item.");

        Assert.NotEmpty(thread!.OtherItems);

        foreach (var item in thread.OtherItems)
        {
            Assert.False(string.IsNullOrWhiteSpace(item.ItemType), "A thread item arrived with no type.");
            Assert.False(
                int.TryParse(item.ItemType, out _),
                $"Thread item type '{item.ItemType}' is a raw class ordinal rather than a name.");
            Assert.False(
                string.IsNullOrWhiteSpace(item.FolderPath),
                $"Thread item '{item.Subject}' did not say which folder it lives in.");
        }

        // Everything the thread holds is now either a message or an identified item. The skipped
        // counter is for entries that genuinely could not be read, which is a different failure.
        Assert.Equal(
            thread.TotalItemCount,
            thread.Messages.Count + thread.OtherItems.Count + thread.SkippedItemCount);

        output.WriteLine(
            $"'{thread.ConversationTopic}': {thread.Messages.Count} message(s), "
            + $"other: {string.Join(", ", thread.OtherItems.Select(i => i.ItemType))}");
    }

    /// <summary>
    /// <c>skippedItemCount</c> must mean "this could not be read" and nothing else. While it also
    /// counted every meeting item, the two were indistinguishable: a thread that was read perfectly
    /// and a thread with unreachable entries produced the same non-zero number, so neither was
    /// actionable.
    /// </summary>
    [SkippableFact]
    public void GetConversation_DoesNotCountReadableNonMailItemsAsSkipped()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var thread = FindThreadWithNonMailItems(commands);

        Skip.If(thread is null, "No conversation in this inbox contains a non-mail item.");

        Assert.Equal(0, thread!.SkippedItemCount);
    }

    /// <summary>
    /// Walks the inbox for a conversation that contains something other than mail. Written not to
    /// assume any particular message, because a test that only passes on one mailbox stops testing
    /// anything on every other one.
    ///
    /// <para>
    /// The selection criterion is deliberately <em>not</em> "has other items" - that is the thing
    /// under test, so searching by it makes the test skip rather than fail when the behaviour is
    /// removed, and a skipped test is a green test. It selects on a thread holding more entries than
    /// it returned messages, which is true whether or not those entries are reported.
    /// </para>
    /// </summary>
    private static MailConversationResult? FindThreadWithNonMailItems(MailCommands commands)
    {
        var listing = commands.List(folder: "Inbox", maxCount: 25);
        if (!listing.Success)
        {
            return null;
        }

        foreach (var message in listing.Messages)
        {
            if (string.IsNullOrWhiteSpace(message.EntryId))
            {
                continue;
            }

            var thread = commands.GetConversation(
                entryId: message.EntryId,
                storeId: message.StoreId,
                useActiveMail: false,
                maxCount: 500);

            if (thread.Success && !thread.Truncated && thread.TotalItemCount > thread.Messages.Count)
            {
                return thread;
            }
        }

        return null;
    }

    /// <summary>
    /// The identifier on a read result and the identifier the thread call reports must agree,
    /// otherwise a caller cannot get from "this message" to "its thread".
    /// </summary>
    [SkippableFact]
    public void Read_ExposesTheSameConversationIdentifiersAsGetConversation()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? originalId = null;

        try
        {
            originalId = CreateDraft(commands, $"OutlookMcp ids {Guid.NewGuid():N}");

            var read = commands.Read(entryId: originalId, useActiveMail: false);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.False(string.IsNullOrWhiteSpace(read.ConversationId));

            var thread = commands.GetConversation(entryId: originalId, useActiveMail: false);
            Assert.True(thread.Success, thread.ErrorMessage);

            Assert.Equal(read.ConversationId, thread.ConversationId);
            Assert.Equal(read.ConversationTopic, thread.ConversationTopic);
        }
        finally
        {
            DeleteQuietly(commands, originalId);
        }
    }

    /// <summary>
    /// A listing has to carry the conversation id too, otherwise reaching a thread costs an extra
    /// read per message - which for an agent scanning a folder is the difference between one call
    /// and twenty-five.
    /// </summary>
    [SkippableFact]
    public void MailList_CarriesConversationIdentifiers()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? originalId = null;

        try
        {
            originalId = CreateDraft(commands, $"OutlookMcp list {Guid.NewGuid():N}");

            var listed = commands.List(folder: "drafts", maxCount: 25);
            Assert.True(listed.Success, listed.ErrorMessage);

            var mine = listed.Messages.FirstOrDefault(m => m.EntryId == originalId);
            Assert.NotNull(mine);
            Assert.False(string.IsNullOrWhiteSpace(mine!.ConversationId));
        }
        finally
        {
            DeleteQuietly(commands, originalId);
        }
    }

    /// <summary>
    /// An unresolvable id must fail loudly. An empty thread reported as success is the
    /// "confidently wrong answer" failure this project keeps finding.
    /// </summary>
    [SkippableFact]
    public void GetConversation_WithAnUnknownEntryId_FailsExplicitly()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        var thread = commands.GetConversation(entryId: "00000000DEADBEEF", useActiveMail: false);

        Assert.False(thread.Success);
        Assert.False(string.IsNullOrWhiteSpace(thread.ErrorMessage));
        Assert.Empty(thread.Messages);
    }

    /// <summary>
    /// maxCount must bound the thread and say so, rather than silently returning a partial thread a
    /// caller would read as the whole conversation.
    /// </summary>
    [SkippableFact]
    public void GetConversation_HonoursMaxCountAndReportsTruncation()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? replyId = null;

        try
        {
            string originalId = FindReceivedMessage(commands);
            replyId = CreateReply(commands, originalId);

            var full = ReadThreadContaining(commands, originalId, replyId);
            Assert.True(full.Success, full.ErrorMessage);
            Skip.If(full.ReturnedCount < 2, "The reply draft did not join the original's conversation.");

            var capped = commands.GetConversation(entryId: originalId, useActiveMail: false, maxCount: 1);
            Assert.True(capped.Success, capped.ErrorMessage);

            Assert.Single(capped.Messages);
            Assert.True(capped.Truncated);
            Assert.True(capped.TotalItemCount >= 2);
        }
        finally
        {
            DeleteQuietly(commands, replyId);
        }
    }

    private static readonly string[] ReceivedMessageFolders = ["inbox", "Inbox/older", "Archive"];

    /// <summary>
    /// Finds a real received message to thread from. More than one folder is searched: an Inbox that
    /// happens to be empty would otherwise make every one of these tests skip, which looks like a
    /// pass and verifies nothing.
    /// </summary>
    private static string FindReceivedMessage(MailCommands commands)
    {
        foreach (string folder in ReceivedMessageFolders)
        {
            var listed = commands.List(folder: folder, maxCount: 25);

            if (!listed.Success)
            {
                continue;
            }

            var candidate = listed.Messages.FirstOrDefault(
                m => !m.IsDraft
                     && !string.IsNullOrWhiteSpace(m.EntryId)
                     && (m.ItemType == null || m.ItemType == "mail"));

            if (candidate != null)
            {
                return candidate.EntryId!;
            }
        }

        throw new SkipException("This mailbox holds no received message to build a thread from.");
    }

    private static string CreateDraft(MailCommands commands, string marker)
    {
        var draft = commands.CreateMailDraft(subject: marker, body: $"{marker}\r\nCreated by an OutlookMcp integration test.");
        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));
        return draft.EntryId!;
    }

    private static string CreateReply(MailCommands commands, string entryId)
    {
        var reply = commands.Reply(entryId: entryId, useActiveMail: false, body: "Reply created by an OutlookMcp integration test.");
        Assert.True(reply.Success, reply.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(reply.EntryId));
        return reply.EntryId!;
    }

    /// <summary>
    /// Reads the thread, waiting for a newly created item to appear in it.
    ///
    /// <para>
    /// Outlook's conversation index is eventually consistent: a reply draft saved a moment ago is
    /// already part of the conversation, but <c>Conversation.GetTable()</c> may not list it yet.
    /// Asserting immediately passes or fails depending on how busy the machine is - it passed when
    /// this test ran alone and failed when it ran after fifty others.
    /// </para>
    ///
    /// <para>
    /// The wait is bounded and the failure is loud. "Eventually" is the real behaviour; "never" is
    /// still a bug, and this must not degrade into waiting until the assertion happens to hold.
    /// </para>
    /// </summary>
    private static MailConversationResult ReadThreadContaining(MailCommands commands, string entryId, string expectedEntryId)
    {
        MailConversationResult thread = null!;

        for (int attempt = 0; attempt < 20; attempt++)
        {
            thread = commands.GetConversation(entryId: entryId, useActiveMail: false);
            Assert.True(thread.Success, thread.ErrorMessage);

            if (thread.Messages.Any(m => m.EntryId == expectedEntryId))
            {
                return thread;
            }

            Thread.Sleep(500);
        }

        return thread;
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
            output.WriteLine("Skipping Outlook conversation test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInterop.NameSpace? session = null;

        try
        {
            session = application.GetNamespace("MAPI");
            _ = session.Folders.Count;
        }
        catch (Exception ex)
        {
            output.WriteLine($"Skipping Outlook conversation test: {ex.Message}");
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            if (session != null && Marshal.IsComObject(session))
            {
                _ = Marshal.FinalReleaseComObject(session);
            }

            if (Marshal.IsComObject(application))
            {
                _ = Marshal.ReleaseComObject(application);
            }
        }
    }
}
