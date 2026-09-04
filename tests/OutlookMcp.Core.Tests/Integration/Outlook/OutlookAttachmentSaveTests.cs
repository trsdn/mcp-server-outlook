using OutlookMcp.Core.Commands.Attachment;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// attachment.save selection semantics (#15). The point of this slice is the off-by-one boundary:
/// Outlook's <c>Attachments</c> collection is 1-based, and the old surface overloaded
/// <c>attachmentIndex = 0</c> to mean "save every attachment". An LLM that assumed 0-based indexing
/// and passed <c>0</c> silently got a completely different operation than it intended. That is the
/// bug. These tests pin the new contract: a 1-based index with no sentinel, an explicit
/// <c>allAttachments</c> flag, selection by <c>attachmentName</c>, and a loud error for <c>0</c>.
///
/// <para>
/// Every test builds its own draft with two real, distinctly named attachments, so nothing here
/// depends on finding suitable mail in the user's mailbox, and nothing here touches a message it did
/// not create. The draft and the files it writes are deleted in a <c>finally</c>.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Attachment")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookAttachmentSaveTests(ITestOutputHelper output)
{
    private const string Marker = "OutlookMcp attach-save test";

    [SkippableFact]
    public void Save_ByIndexOne_SavesTheFirstAttachment()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, string firstName, _) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentIndex: 1,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Null(result.ErrorMessage);
            Assert.Equal(1, result.SavedCount);
            Assert.True(File.Exists(Path.Combine(dest, firstName)), "The first attachment was not written.");
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    [SkippableFact]
    public void Save_ByLastValidIndex_SavesTheLastAttachment()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, string secondName) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentIndex: 2,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(1, result.SavedCount);
            Assert.True(File.Exists(Path.Combine(dest, secondName)), "The last attachment was not written.");
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    /// <summary>
    /// The heart of the fix. <c>0</c> used to mean "all"; against 1-based COM it is not a valid index
    /// and must now be a clear error, never a silent "save everything" and never "save attachment 1".
    /// </summary>
    [SkippableFact]
    public void Save_ByIndexZero_IsRejected_AndDoesNotSaveEverything()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, _) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentIndex: 0,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            // It must point the caller at the explicit way to say "all".
            Assert.Contains("allAttachments", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            // Nothing may have been written: 0 is not "all" any more.
            Assert.Equal(0, result.SavedCount);
            Assert.Empty(Directory.GetFiles(dest));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    [SkippableFact]
    public void Save_ByNegativeIndex_IsRejected()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, _) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentIndex: -1,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Empty(Directory.GetFiles(dest));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    /// <summary>One past the end, with a message that names the valid range.</summary>
    [SkippableFact]
    public void Save_ByIndexPastEnd_IsRejected_WithRange()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, _) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentIndex: 3,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Contains("1 and 2", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(Directory.GetFiles(dest));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    /// <summary>The explicit replacement for the former magic 0.</summary>
    [SkippableFact]
    public void Save_AllAttachments_SavesEveryAttachment()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, string firstName, string secondName) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                allAttachments: true,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(2, result.SavedCount);
            Assert.True(File.Exists(Path.Combine(dest, firstName)));
            Assert.True(File.Exists(Path.Combine(dest, secondName)));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    /// <summary>Names are what the model actually saw in attachment.list; they are the better key.</summary>
    [SkippableFact]
    public void Save_ByName_SavesThatAttachment()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, string secondName) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentName: secondName,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(1, result.SavedCount);
            Assert.True(File.Exists(Path.Combine(dest, secondName)));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    [SkippableFact]
    public void Save_ByUnknownName_IsRejected()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, _) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentName: "no-such-file-" + Guid.NewGuid().ToString("N") + ".bin",
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Empty(Directory.GetFiles(dest));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    /// <summary>Conflicting selectors are ambiguous and must be refused rather than silently ranked.</summary>
    [SkippableFact]
    public void Save_WithConflictingSelectors_IsRejected()
    {
        EnsureOutlookAvailable();

        using var work = new Workspace();
        var attachments = new AttachmentCommands();
        var mail = new MailCommands();
        string? draftId = null;

        try
        {
            (draftId, _, _) = CreateDraftWithTwoAttachments(work, attachments, mail);
            string dest = work.NewOutputDir();

            var result = attachments.Save(
                destinationDirectory: dest,
                attachmentIndex: 1,
                allAttachments: true,
                mailEntryId: draftId,
                useActiveMail: false);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Empty(Directory.GetFiles(dest));
        }
        finally
        {
            DeleteQuietly(mail, draftId);
        }
    }

    private static (string DraftId, string FirstName, string SecondName) CreateDraftWithTwoAttachments(
        Workspace work, AttachmentCommands attachments, MailCommands mail)
    {
        var draft = mail.CreateMailDraft(subject: Marker, body: "placeholder");
        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));
        string draftId = draft.EntryId!;

        string first = work.NewInputFile("first", "first attachment payload");
        string second = work.NewInputFile("second", "second attachment payload");

        var add1 = attachments.Add(filePath: first, mailEntryId: draftId, useActiveMail: false);
        Assert.True(add1.Success, add1.ErrorMessage);
        var add2 = attachments.Add(filePath: second, mailEntryId: draftId, useActiveMail: false);
        Assert.True(add2.Success, add2.ErrorMessage);

        return (draftId, Path.GetFileName(first), Path.GetFileName(second));
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
            output.WriteLine("Skipping Outlook attachment-save test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }

    /// <summary>
    /// A per-test scratch area under the test's own output directory (never a system temp path),
    /// holding the files we attach and the directories we save into. Everything is removed on dispose.
    /// </summary>
    private sealed class Workspace : IDisposable
    {
        private readonly string _root = Path.Combine(
            AppContext.BaseDirectory, "attach-save-tests", Guid.NewGuid().ToString("N"));

        private readonly string _prefix = "omcp-" + Guid.NewGuid().ToString("N")[..8];
        private int _counter;

        public Workspace() => Directory.CreateDirectory(_root);

        public string NewInputFile(string label, string content)
        {
            string path = Path.Combine(_root, $"{_prefix}-{label}-{_counter++}.txt");
            File.WriteAllText(path, content);
            return path;
        }

        public string NewOutputDir()
        {
            string path = Path.Combine(_root, "out-" + _counter++);
            Directory.CreateDirectory(path);
            return path;
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_root))
                {
                    Directory.Delete(_root, recursive: true);
                }
            }
            catch (IOException)
            {
                // Best-effort cleanup of scratch files; a lingering lock must not fail the test.
            }
        }
    }
}
