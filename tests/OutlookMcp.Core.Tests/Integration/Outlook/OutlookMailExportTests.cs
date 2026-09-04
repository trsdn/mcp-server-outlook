using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Item export - <c>MailItem.SaveAs</c> / <c>AppointmentItem.SaveAs</c> (#14).
///
/// <para>
/// <b>Every rule asserted here was measured against real Outlook first.</b> <c>SaveAs</c> is a
/// deceptively simple method with four separate ways to hand a caller a confidently wrong result,
/// and all four were reproduced before any of this was designed:
/// </para>
///
/// <para>
/// <b>The ANSI .msg format silently destroys text.</b> Saving a subject reading
/// <c>probe Grüsse äöü тест €</c> with <c>olMSG</c> (3) and reading it back gives
/// <c>probe Grüsse äöü ???? €</c>. No error, no warning - the Cyrillic is simply gone.
/// <c>olMSGUnicode</c> (9) round-trips it exactly. So <c>msg</c> always means Unicode here, and
/// this test is what stops anyone "simplifying" that back to the constant whose name matches.
/// </para>
///
/// <para>
/// <b>The extension is ignored.</b> <c>SaveAs("x.txt", olMSGUnicode)</c> succeeds and writes an OLE
/// compound file - first bytes <c>D0 CF 11 E0</c> - to a file called <c>.txt</c>. Nothing errors.
/// </para>
///
/// <para>
/// <b>A relative path succeeds and lands somewhere else.</b> It resolves against Outlook's own
/// working directory, not the caller's, so the file exists but is not where anyone will look.
/// </para>
///
/// <para>
/// <b>A missing directory reports "The operation failed."</b> - which tells the caller nothing.
/// </para>
///
/// <para>
/// <b>Mutation safety.</b> Every test exports a scratch draft it created itself, deletes it in
/// <c>finally</c>, and writes only inside a per-test temp directory that is removed afterwards. No
/// pre-existing item is touched, and nothing is written outside the temp tree.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Export")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMailExportTests(ITestOutputHelper output)
{
    /// <summary>
    /// The whole reason <c>msg</c> maps to <c>olMSGUnicode</c>: exporting and reopening must return
    /// the subject that went in, including characters outside the machine's ANSI code page. With
    /// <c>olMSG</c> the Cyrillic comes back as <c>????</c> and nothing reports a problem.
    /// </summary>
    [SkippableFact]
    public void Export_ToMsg_RoundTripsTextThatAnsiWouldDestroy()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        using var scratch = new ScratchDirectory();
        string subject = ScratchSubject() + " Grüsse äöü тест €";
        string? entryId = null;
        string? storeId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: subject, body: "export round-trip probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            string path = Path.Combine(scratch.Path, "roundtrip.msg");
            var exported = commands.Export(path, entryId: entryId, storeId: storeId, useActiveMail: false);

            Assert.True(exported.Success, exported.ErrorMessage);
            Assert.Equal("msg", exported.Format);
            Assert.True(File.Exists(path), $"Export reported success but {path} does not exist.");
            Assert.True(exported.BytesWritten > 0);
            Assert.Equal(new FileInfo(path).Length, exported.BytesWritten);

            // An OLE compound file, which is what a .msg is. If msg were ever mapped to a text
            // format this signature would change and the assertion below would still pass, so both
            // are checked.
            byte[] header = new byte[4];
            using (var stream = File.OpenRead(path))
            {
                Assert.Equal(4, stream.Read(header, 0, 4));
            }
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, header);

            string reopened = ReadSubjectFromSavedItem(path);
            Assert.Equal(subject, reopened);

            output.WriteLine($"round-tripped: {reopened} ({exported.BytesWritten} bytes)");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// A relative path must be refused. Outlook accepts one and resolves it against its own working
    /// directory, so the caller gets success and a file they will never find.
    /// </summary>
    [SkippableFact]
    public void Export_RefusesARelativePathRatherThanWritingItSomewhereElse()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: ScratchSubject(), body: "relative path probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            var exported = commands.Export(
                "relative-export-probe.msg",
                entryId: entryId,
                storeId: storeId,
                useActiveMail: false);

            Assert.False(exported.Success);
            Assert.False(string.IsNullOrWhiteSpace(exported.ErrorMessage));
            Assert.Contains("absolute", exported.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// A missing directory must be named. Outlook's own answer is "The operation failed.", which
    /// gives a caller nothing to act on.
    /// </summary>
    [SkippableFact]
    public void Export_NamesTheMissingDirectoryInsteadOfOutlooksBlankFailure()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        using var scratch = new ScratchDirectory();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: ScratchSubject(), body: "missing dir probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            string missing = Path.Combine(scratch.Path, "no-such-directory");
            var exported = commands.Export(
                Path.Combine(missing, "x.msg"),
                entryId: entryId,
                storeId: storeId,
                useActiveMail: false);

            Assert.False(exported.Success);
            Assert.Contains("no-such-directory", exported.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// An existing file is not replaced unless the caller asks. Outlook overwrites silently, and an
    /// export that quietly destroys the previous export is not recoverable.
    /// </summary>
    [SkippableFact]
    public void Export_RefusesToOverwriteUnlessAsked()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        using var scratch = new ScratchDirectory();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: ScratchSubject(), body: "overwrite probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            string path = Path.Combine(scratch.Path, "twice.msg");
            var first = commands.Export(path, entryId: entryId, storeId: storeId, useActiveMail: false);
            Assert.True(first.Success, first.ErrorMessage);
            Assert.False(first.Overwritten);

            var second = commands.Export(path, entryId: entryId, storeId: storeId, useActiveMail: false);
            Assert.False(second.Success);
            Assert.Contains("overwrite", second.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            var forced = commands.Export(
                path,
                entryId: entryId,
                storeId: storeId,
                useActiveMail: false,
                overwrite: true);
            Assert.True(forced.Success, forced.ErrorMessage);
            Assert.True(forced.Overwritten);

            output.WriteLine($"refused, then overwrote on request: {forced.FilePath}");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// If the caller names a format and an extension that disagree, refuse. Outlook writes the
    /// requested format regardless of the extension, so the alternative is a binary .msg living
    /// under a .txt name - a file that every other tool will misread.
    /// </summary>
    [SkippableFact]
    public void Export_RefusesAFormatThatContradictsTheExtension()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        using var scratch = new ScratchDirectory();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: ScratchSubject(), body: "mismatch probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            string path = Path.Combine(scratch.Path, "actually-msg.txt");
            var exported = commands.Export(
                path,
                format: "msg",
                entryId: entryId,
                storeId: storeId,
                useActiveMail: false);

            Assert.False(exported.Success);
            Assert.False(File.Exists(path), "A rejected export still wrote a file.");
            Assert.Contains("txt", exported.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// The format is taken from the extension when none is given, and a text export really is text -
    /// so the "extension is ignored" trap cannot come back through the default path either.
    /// </summary>
    [SkippableFact]
    public void Export_DerivesTheFormatFromTheExtension()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        using var scratch = new ScratchDirectory();
        string? entryId = null;
        string? storeId = null;

        try
        {
            string subject = ScratchSubject();
            var draft = commands.CreateMailDraft(subject: subject, body: "derived format probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            string path = Path.Combine(scratch.Path, "derived.txt");
            var exported = commands.Export(path, entryId: entryId, storeId: storeId, useActiveMail: false);

            Assert.True(exported.Success, exported.ErrorMessage);
            Assert.Equal("txt", exported.Format);

            string text = File.ReadAllText(path);
            Assert.Contains(subject, text, StringComparison.Ordinal);

            output.WriteLine($"derived txt from extension, {exported.BytesWritten} bytes");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// A format the item cannot produce must be named. Outlook answers a mail item asked for iCal
    /// with "Value does not fall within the expected range", which reads like a bug in the caller's
    /// arguments rather than "mail is not a calendar entry".
    /// </summary>
    [SkippableFact]
    public void Export_ExplainsWhyMailCannotBeSavedAsICalendar()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        using var scratch = new ScratchDirectory();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var draft = commands.CreateMailDraft(subject: ScratchSubject(), body: "ics probe");
            Assert.True(draft.Success, draft.ErrorMessage);
            entryId = draft.EntryId;
            storeId = draft.StoreId;

            var exported = commands.Export(
                Path.Combine(scratch.Path, "mail.ics"),
                entryId: entryId,
                storeId: storeId,
                useActiveMail: false);

            Assert.False(exported.Success);
            Assert.False(string.IsNullOrWhiteSpace(exported.ErrorMessage));
            Assert.DoesNotContain(
                "expected range",
                exported.ErrorMessage!,
                StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// An unresolvable id must not produce a file. An export that writes an empty or wrong item and
    /// reports success is worse than one that fails.
    /// </summary>
    [SkippableFact]
    public void Export_RefusesAnIdThatDoesNotResolveAndWritesNothing()
    {
        EnsureOutlookAvailable();

        using var scratch = new ScratchDirectory();
        string path = Path.Combine(scratch.Path, "never.msg");

        var exported = new MailCommands().Export(
            path,
            entryId: "0000000000000000000000000000000000000000000000",
            useActiveMail: false);

        Assert.False(exported.Success);
        Assert.False(File.Exists(path), "A failed export still wrote a file.");

        output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
    }

    private static string ReadSubjectFromSavedItem(string path)
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            throw new InvalidOperationException("Outlook went away mid-test.");
        }

        OutlookInterop.NameSpace? session = null;
        object? item = null;
        try
        {
            session = application!.GetNamespace("MAPI");
            item = session.OpenSharedItem(path);
            string subject = ((dynamic)item).Subject?.ToString() ?? string.Empty;
            ((dynamic)item).Close(OutlookInterop.OlInspectorClose.olDiscard);
            return subject;
        }
        finally
        {
            if (item != null)
            {
                OutlookInteropRunner.ReleaseComObject(ref item!);
            }

            if (session != null)
            {
                OutlookInteropRunner.ReleaseComObject(ref session!);
            }

            // Plain decrement, never FinalReleaseComObject: this is the shared Outlook.Application
            // and final-releasing it breaks every other holder in the process (#19, #116).
            OutlookInteropRunner.ReleaseSharedComObject(ref application);
        }
    }

    private void DeleteIfCreated(MailCommands commands, string? entryId, string? storeId)
    {
        if (string.IsNullOrEmpty(entryId))
        {
            return;
        }

        var deleted = commands.Delete(entryId, storeId);
        output.WriteLine($"Delete draft: success={deleted.Success} {deleted.ErrorMessage}");
    }

    private const string ScratchPrefix = "mcp-export-test-";

    private static string ScratchSubject() => $"{ScratchPrefix}{Guid.NewGuid():N}";

    /// <summary>
    /// A temp directory that is removed even when a test fails, so a failed run cannot leave export
    /// files behind.
    /// </summary>
    private sealed class ScratchDirectory : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"mcp-export-{Guid.NewGuid():N}");

        public ScratchDirectory() => Directory.CreateDirectory(Path);

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Path))
                {
                    Directory.Delete(Path, recursive: true);
                }
            }
            catch (IOException)
            {
                // A file still held open by Outlook is not worth failing a test over; the temp
                // directory is per-run and named, so a leftover is identifiable.
            }
        }
    }

    private static void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        // Plain decrement, never FinalReleaseComObject: this is the shared Outlook.Application and
        // final-releasing it breaks every other holder in the process (#19, #116).
        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
