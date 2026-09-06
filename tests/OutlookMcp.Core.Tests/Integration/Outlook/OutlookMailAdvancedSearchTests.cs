using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Free-text search answered by <c>Application.AdvancedSearch</c> (#13, #27).
///
/// <para>
/// <c>Items.Restrict</c> cannot filter on <c>Body</c> at all, so until now a body match meant opening
/// every candidate item and giving up at a scan limit - which is to say that a term in a message
/// further back than the limit was reported as "no such mail". <c>AdvancedSearch</c> asks Outlook to
/// run the same substring question with no client-side scan and no horizon.
/// </para>
///
/// <para>
/// The tests that matter here are the ones that distinguish the engines rather than merely observing
/// that a search returned something: which engine answered, that a body-only term is still found, and
/// that substring matching survived the move (the content index would silently turn it into whole-word
/// matching, and the caller would never see what it stopped finding).
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailAdvancedSearch")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailAdvancedSearchTests(ITestOutputHelper output)
{
    private const string Marker = "OutlookMcp-advsearch";

    /// <summary>
    /// The change itself. Reported rather than inferred, because an empty result means something
    /// different depending on which engine produced it.
    /// </summary>
    [SkippableFact]
    public void MailSearch_ByDefault_IsAnsweredByAdvancedSearch()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string token = UniqueToken();

        try
        {
            draftId = CreateDraftWithBodyToken(commands, token);

            var found = commands.Search(query: token, folder: "drafts", maxCount: 25);

            Assert.True(found.Success, found.ErrorMessage);
            output.WriteLine($"engine={found.SearchEngine} projection={found.Projection} returned={found.ReturnedCount} message={found.Message}");
            Assert.Equal("advancedSearch", found.SearchEngine);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// The reason the engine exists: the term is nowhere but the body, which is the one place
    /// <c>Restrict</c> cannot look.
    /// </summary>
    [SkippableFact]
    public void MailSearch_ByDefault_FindsATermThatAppearsOnlyInTheBody()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string token = UniqueToken();

        try
        {
            draftId = CreateDraftWithBodyToken(commands, token);

            var found = commands.Search(query: token, folder: "drafts", maxCount: 25);

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Contains(found.Messages, m => m.EntryId == draftId);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// Substring matching has to survive the move.
    ///
    /// <para>
    /// The content index matches whole words: it finds <c>foo</c> in "a foo arrived" but not inside
    /// <c>foobar</c>. If the default search quietly became an indexed one, every mid-word match would
    /// stop being found and the caller would be told the mail does not exist. This searches for a
    /// token that appears only in the middle of a longer word.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void MailSearch_ByDefault_StillMatchesInsideALongerWord()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string token = UniqueToken();

        try
        {
            var draft = commands.CreateMailDraft(
                subject: UniqueSubject(),
                body: "prefix" + token + "suffix");
            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            var found = commands.Search(query: token, folder: "drafts", maxCount: 25);

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Contains(found.Messages, m => m.EntryId == draftId);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// A search answered by Outlook does not open the messages it returns, so it gets the same cheap
    /// projection an ordinary listing does.
    /// </summary>
    [SkippableFact]
    public void MailSearch_ByDefault_DoesNotOpenTheMessagesItReturns()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string token = UniqueToken();

        try
        {
            draftId = CreateDraftWithBodyToken(commands, token);

            var found = commands.Search(query: token, folder: "drafts", maxCount: 25);

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Equal("table", found.Projection);

            // And the rows must still be usable: an id that does not round-trip would make the whole
            // result set decorative.
            var mine = found.Messages.FirstOrDefault(m => m.EntryId == draftId);
            Assert.NotNull(mine);

            var read = commands.Read(entryId: mine!.EntryId, storeId: mine.StoreId, useActiveMail: false);
            Assert.True(read.Success, read.ErrorMessage);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// Structured filters are not abandoned when the free-text engine changes: the DASL predicates go
    /// into the same search, so "this word, from this folder, unread" stays one call.
    /// </summary>
    [SkippableFact]
    public void MailSearch_ByDefault_StillAppliesStructuredFilters()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? matching = null;
        string? excluded = null;
        string token = UniqueToken();

        try
        {
            matching = CreateDraftWithBodyToken(commands, token);
            excluded = CreateDraftWithBodyToken(commands, token);

            var flagged = commands.SetFlag(entryId: matching, useActiveMail: false);
            Assert.True(flagged.Success, flagged.ErrorMessage);

            var found = commands.Search(query: token, folder: "drafts", maxCount: 25, flaggedOnly: true);

            Assert.True(found.Success, found.ErrorMessage);
            Assert.Contains(found.Messages, m => m.EntryId == matching);
            Assert.DoesNotContain(found.Messages, m => m.EntryId == excluded);
        }
        finally
        {
            DeleteQuietly(commands, matching);
            DeleteQuietly(commands, excluded);
        }
    }

    /// <summary>
    /// The other two engines stay reachable and keep saying which one answered. A caller that needs
    /// the bounded, item-by-item scan must still be able to ask for it by name.
    /// </summary>
    [SkippableTheory]
    [InlineData("clientScan", "clientScan")]
    [InlineData("fullText", "contentIndex")]
    [InlineData("advancedSearch", "advancedSearch")]
    public void MailSearch_WithAnExplicitMode_UsesThatEngine(string mode, string expectedEngine)
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string token = UniqueToken();

        try
        {
            draftId = CreateDraftWithBodyToken(commands, token);

            var found = commands.Search(query: token, folder: "drafts", maxCount: 25, searchMode: mode);

            Assert.True(found.Success, found.ErrorMessage);
            output.WriteLine($"mode={mode} engine={found.SearchEngine} message={found.Message}");
            Assert.Equal(expectedEngine, found.SearchEngine);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    /// <summary>
    /// An unrecognised mode is still refused rather than defaulted. Adding an engine must not turn a
    /// typo into "you got the default and nobody told you".
    /// </summary>
    [SkippableFact]
    public void MailSearch_WithAnUnknownSearchMode_IsRefusedAndNamesTheOptions()
    {
        EnsureOutlookAvailable();

        var refused = new MailCommands().Search(query: "anything", folder: "drafts", searchMode: "turbo");

        Assert.False(refused.Success);
        Assert.Contains("turbo", refused.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("advancedSearch", refused.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A wildcard in the query cannot be escaped in DASL, so it cannot be pushed down without
    /// silently widening the search. It must fall back to an engine that can answer it exactly rather
    /// than returning a different result set under the same label.
    /// </summary>
    [SkippableFact]
    public void MailSearch_WithAWildcardInTheQuery_FallsBackRatherThanWidening()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? draftId = null;
        string token = UniqueToken();

        try
        {
            var draft = commands.CreateMailDraft(subject: UniqueSubject(), body: "a 100% certain " + token);
            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            var found = commands.Search(query: "100% certain " + token, folder: "drafts", maxCount: 25);

            Assert.True(found.Success, found.ErrorMessage);
            output.WriteLine($"engine={found.SearchEngine} message={found.Message}");
            Assert.Equal("clientScan", found.SearchEngine);
            Assert.Contains(found.Messages, m => m.EntryId == draftId);
        }
        finally
        {
            DeleteQuietly(commands, draftId);
        }
    }

    private static string UniqueToken() => "advtok" + Guid.NewGuid().ToString("N")[..12];

    private static string UniqueSubject() => $"{Marker} {Guid.NewGuid():N}";

    private static string CreateDraftWithBodyToken(MailCommands commands, string token)
    {
        var draft = commands.CreateMailDraft(
            subject: UniqueSubject(),
            body: "carrier text " + token + " trailing text");

        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));
        return draft.EntryId!;
    }

    /// <summary>
    /// Deletes a scratch draft this test created, and <b>fails the test</b> if it could not.
    ///
    /// <para>
    /// This used to discard the result, which meant a delete that failed because Outlook was briefly
    /// unresponsive (#139) left the draft in the user's real Drafts folder and still reported green.
    /// That happened: a draft from a run at 01:46 was found still sitting in the live mailbox. The
    /// same trade-off applies as in the sibling projection tests - this asserts from a
    /// <c>finally</c>, so a body that also threw has its exception replaced, which is accepted
    /// because the previously invisible case (body passed, residue left behind) is the one worth
    /// catching and the message names cleanup as the cause.
    /// </para>
    /// </summary>
    private static void DeleteQuietly(MailCommands commands, string? entryId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            return;
        }

        var deleted = commands.Delete(entryId: entryId, useActiveMail: false);

        if (!deleted.Success)
        {
            System.Threading.Thread.Sleep(500);
            deleted = commands.Delete(entryId: entryId, useActiveMail: false);
        }

        Assert.True(
            deleted.Success,
            $"Cleanup failed: scratch draft '{entryId}' could not be deleted and is now left in the "
            + $"user's real Drafts folder. Outlook reported: {deleted.ErrorMessage}");
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping Outlook advanced-search test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
