using System.Globalization;
using System.Text.Json;
using OutlookMcp.ComInterop.ServiceClient;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Follow-up flags (#15).
///
/// <para>
/// A flag is only worth setting if something can find it again, so these cover both halves: writing
/// the flag, and seeing it come back on <c>read</c> and <c>list</c>. A write-only flag would be a
/// feature an agent could use once and never act on.
/// </para>
///
/// <para>
/// Flag state is read back through <b>raw COM</b> as well as through this project's reader. Checking
/// only our own reader would pass if the write and the read agreed on the wrong thing, which is the
/// failure this repository keeps finding.
/// </para>
///
/// <para>
/// Nothing is sent, and nothing belonging to the user is touched: every message here is a draft
/// these tests created and delete again in a <c>finally</c>.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailFlag")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailFlagTests(ITestOutputHelper output)
{
    private const string Marker = "OutlookMcp flag test";

    // Outlook.OlFlagStatus
    private const int NoFlag = 0;
    private const int FlagMarked = 2;

    [SkippableFact]
    public void SetFlag_MarksTheMessageForFollowUp()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            Assert.Equal(NoFlag, ReadFlagStatus(draftId));

            var flagged = commands.SetFlag(entryId: draftId, useActiveMail: false);

            Assert.True(flagged.Success, flagged.ErrorMessage);
            Assert.Equal(FlagMarked, ReadFlagStatus(draftId));
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// A due date is the only part of a flag that makes "what is overdue" answerable, so it has to
    /// survive as a date rather than as text in the label.
    /// </summary>
    [SkippableFact]
    public void SetFlag_WithADueDate_StoresItAsADate()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            DateTime due = DateTime.Today.AddDays(3);

            var flagged = commands.SetFlag(
                entryId: draftId,
                useActiveMail: false,
                dueDate: due.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));

            Assert.True(flagged.Success, flagged.ErrorMessage);

            DateTime? stored = ReadTaskDueDate(draftId);

            Assert.NotNull(stored);
            Assert.Equal(due.Date, stored!.Value.Date);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The label an agent sets has to be the label a human sees, or "Review before Friday" silently
    /// becomes a generic follow-up.
    /// </summary>
    [SkippableFact]
    public void SetFlag_WithAFlagRequest_UsesItAsTheLabel()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            var flagged = commands.SetFlag(
                entryId: draftId,
                useActiveMail: false,
                flagRequest: "Review before Friday");

            Assert.True(flagged.Success, flagged.ErrorMessage);
            Assert.Equal("Review before Friday", ReadFlagRequest(draftId));
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// Outlook refuses to complete a follow-up on a draft - measured, not assumed: assigning
    /// <c>olFlagComplete</c> raises "The object does not support this method". The contract is that
    /// this surfaces as an explanation the caller can act on, not as a raw COM error, and that
    /// nothing is half-applied.
    /// </summary>
    [SkippableFact]
    public void SetFlag_Complete_OnADraft_IsRefusedWithAnExplanation()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            _ = commands.SetFlag(entryId: draftId, useActiveMail: false);

            var done = commands.SetFlag(entryId: draftId, useActiveMail: false, flagStatus: "complete");

            Assert.False(done.Success);
            Assert.NotNull(done.ErrorMessage);
            Assert.Contains("sent or received", done.ErrorMessage, StringComparison.OrdinalIgnoreCase);

            // The flag it already carried must survive a refused completion.
            Assert.Equal(FlagMarked, ReadFlagStatus(draftId));
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// Clearing must also clear the due date. Leaving the old date behind reports an unflagged
    /// message as due on a date nobody set, which reads as a deadline that does not exist.
    /// </summary>
    [SkippableFact]
    public void SetFlag_None_ClearsTheFlagEntirely()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            _ = commands.SetFlag(entryId: draftId, useActiveMail: false, dueDate: DateTime.Today.AddDays(1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            Assert.Equal(FlagMarked, ReadFlagStatus(draftId));

            var cleared = commands.SetFlag(entryId: draftId, useActiveMail: false, flagStatus: "none");

            Assert.True(cleared.Success, cleared.ErrorMessage);
            Assert.Equal(NoFlag, ReadFlagStatus(draftId));
            Assert.Null(ReadTaskDueDate(draftId));

            var read = commands.Read(entryId: draftId, useActiveMail: false);
            Assert.Equal("none", read.FlagStatus);
            Assert.Null(read.FlagDueDate);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// An unknown status must be refused. Quietly defaulting "done" to "flagged" would mark an item
    /// as needing attention at the exact moment the user said it no longer did.
    /// </summary>
    [SkippableFact]
    public void SetFlag_WithAnUnknownStatus_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            var refused = commands.SetFlag(entryId: draftId, useActiveMail: false, flagStatus: "done");

            Assert.False(refused.Success);
            Assert.Contains("done", refused.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);

            // And it must not have half-applied.
            Assert.Equal(NoFlag, ReadFlagStatus(draftId));
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    [SkippableFact]
    public void SetFlag_WithAnUnparseableDueDate_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            var refused = commands.SetFlag(entryId: draftId, useActiveMail: false, dueDate: "next tuesday");

            Assert.False(refused.Success);
            Assert.Equal(NoFlag, ReadFlagStatus(draftId));
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The half that makes the feature usable: having flagged something, an agent has to be able to
    /// find it again.
    /// </summary>
    [SkippableFact]
    public void Read_ReportsTheFlagAndItsDueDate()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            DateTime due = DateTime.Today.AddDays(2);

            _ = commands.SetFlag(
                entryId: draftId,
                useActiveMail: false,
                dueDate: due.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                flagRequest: "Follow up");

            var read = commands.Read(entryId: draftId, useActiveMail: false);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.Equal("flagged", read.FlagStatus);
            Assert.Equal("Follow up", read.FlagRequest);
            Assert.NotNull(read.FlagDueDate);
            Assert.Equal(due.Date, read.FlagDueDate!.Value.Date);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// And from a listing, so "what still needs follow-up" does not cost one read per message.
    /// </summary>
    [SkippableFact]
    public void List_ReportsTheFlagWithoutARoundTripPerMessage()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            _ = commands.SetFlag(entryId: draftId, useActiveMail: false, flagRequest: "Follow up");

            var listed = commands.List(folder: "drafts", maxCount: 100);

            Assert.True(listed.Success, listed.ErrorMessage);

            var mine = listed.Messages.FirstOrDefault(m => m.EntryId == draftId);

            Assert.NotNull(mine);
            output.WriteLine($"listed flagStatus={mine!.FlagStatus} request={mine.FlagRequest}");
            Assert.Equal("flagged", mine.FlagStatus);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// An unflagged message must report "none" rather than omitting the field, so a caller can tell
    /// "not flagged" from "this listing does not report flags".
    /// </summary>
    [SkippableFact]
    public void List_ReportsAnUnflaggedMessageAsNoneRatherThanOmittingIt()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;

        try
        {
            draftId = CreateDraft(commands);

            var listed = commands.List(folder: "drafts", maxCount: 100);
            Assert.True(listed.Success, listed.ErrorMessage);

            var mine = listed.Messages.FirstOrDefault(m => m.EntryId == draftId);

            Assert.NotNull(mine);
            Assert.Equal("none", mine!.FlagStatus);

            // Asserting on the object alone would pass even if the field were dropped from the wire,
            // because "none" is also its default. The contract is that a caller can see it, so this
            // serialises with the same options the transport uses.
            string json = JsonSerializer.Serialize(mine, ServiceProtocol.JsonOptions);
            output.WriteLine(json);
            Assert.Contains("\"flagStatus\"", json, StringComparison.Ordinal);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The filter must find flagged mail, and must not quietly hide it. An under-inclusive Restrict
    /// is the worst failure this surface can have: Outlook drops the item before the client sees it,
    /// so the caller is told the mail does not exist.
    /// </summary>
    [SkippableFact]
    public void List_FlaggedOnly_ReturnsFlaggedMailAndExcludesUnflagged()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? flaggedId = null;
        string? plainId = null;

        try
        {
            flaggedId = CreateDraft(commands);
            plainId = CreateDraft(commands);

            var set = commands.SetFlag(entryId: flaggedId, useActiveMail: false, flagRequest: "Follow up");
            Assert.True(set.Success, set.ErrorMessage);

            var listed = commands.List(folder: "drafts", maxCount: 200, flaggedOnly: true);

            Assert.True(listed.Success, listed.ErrorMessage);
            Assert.Contains(listed.Messages, m => m.EntryId == flaggedId);
            Assert.DoesNotContain(listed.Messages, m => m.EntryId == plainId);

            // Everything it returns must actually be flagged, not merely "was flagged once".
            Assert.All(listed.Messages, m => Assert.Equal("flagged", m.FlagStatus));
        }
        finally
        {
            DeleteQuietly(commands, flaggedId);
            DeleteQuietly(commands, plainId);
        }
    }

    /// <summary>
    /// The filter must be pushed into Outlook rather than applied after hydrating the folder. If it
    /// were client-side, scannedCount would climb to the folder total and the whole point - not
    /// touching every item - would be lost while the results still looked correct.
    /// </summary>
    [SkippableFact]
    public void List_FlaggedOnly_IsEvaluatedByOutlookRatherThanAfterScanningTheFolder()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? flaggedId = null;

        try
        {
            flaggedId = CreateDraft(commands);
            _ = commands.SetFlag(entryId: flaggedId, useActiveMail: false);

            var all = commands.List(folder: "drafts", maxCount: 200);
            var filtered = commands.List(folder: "drafts", maxCount: 200, flaggedOnly: true);

            Assert.True(all.Success, all.ErrorMessage);
            Assert.True(filtered.Success, filtered.ErrorMessage);

            output.WriteLine($"unfiltered scanned={all.ScannedCount} filtered scanned={filtered.ScannedCount}");

            Skip.If(
                all.ScannedCount <= 1,
                "Drafts holds too few items for the scan count to distinguish push-down from a client-side filter.");

            Assert.True(
                filtered.ScannedCount < all.ScannedCount,
                $"Expected the flag filter to be pushed into Outlook, but it scanned {filtered.ScannedCount} " +
                $"of {all.ScannedCount} items - the same work as no filter at all.");
        }
        finally
        {
            DeleteQuietly(commands, flaggedId);
        }
    }

    private static int ReadFlagStatus(string? entryId)
        => WithMail(entryId, mail => (int)mail.FlagStatus);

    private static string? ReadFlagRequest(string? entryId)
        => WithMail(entryId, mail => mail.FlagRequest);

    private static DateTime? ReadTaskDueDate(string? entryId)
        => WithMail(entryId, mail =>
        {
            DateTime due = mail.TaskDueDate;

            // Outlook uses a sentinel far-future date for "no due date".
            return due.Year >= 4000 ? (DateTime?)null : due;
        });

    /// <summary>
    /// Reads a property straight off the <c>MailItem</c> through raw COM, deliberately bypassing this
    /// project's own reader.
    /// </summary>
    private static T WithMail<T>(string? entryId, Func<OutlookInterop.MailItem, T> read)
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

            return read(mail!);
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref mail);
            OutlookInteropRunner.ReleaseComObject(ref session);
            OutlookInteropRunner.ReleaseSharedComObject(ref application);
        }
    }

    private static string CreateDraft(MailCommands commands)
    {
        var draft = commands.CreateMailDraft(subject: $"{Marker} {Guid.NewGuid():N}", body: "placeholder");
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
            output.WriteLine("Skipping Outlook flag test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}



