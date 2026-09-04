using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Application;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Xunit;
using Xunit.Abstractions;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Covers the two UI-context actions recovered from the orphaned parity branch.
///
/// The interesting case is not "an inspector is open" - it is the compose window that has never
/// been saved, which has no entry id and therefore cannot be addressed by any other action in this
/// surface. One of the tests below opens a real compose window to exercise exactly that path,
/// because it is otherwise never reached: on an idle profile there is no inspector at all.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Application")]
[Trait("RequiresOutlook", "true")]
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
public class OutlookUiContextTests(ITestOutputHelper output)
{
    private const string SubjectPrefix = "mcp-uicontext-";

    /// <summary>
    /// An idle profile still has an explorer, and the folder it is looking at is the single most
    /// useful piece of context an agent can have before it starts guessing folder names.
    /// </summary>
    [SkippableFact]
    public void GetActiveExplorer_ReportsTheFolderTheUserIsLookingAt()
    {
        EnsureOutlookAvailable();

        var result = new ApplicationCommands().GetActiveExplorer();

        Assert.True(result.Success, result.ErrorMessage);
        Skip.IfNot(result.HasExplorer, "Outlook is running with no explorer window open.");

        Assert.False(
            string.IsNullOrWhiteSpace(result.CurrentFolderName),
            "An explorer is open but it reported no current folder.");

        output.WriteLine(
            $"Explorer on '{result.CurrentFolderName}' ({result.CurrentFolderPath}), selection={result.SelectionCount}");
    }

    /// <summary>
    /// The selection fields have to agree with each other. Reporting a selected item type while
    /// claiming a selection count of zero - or the reverse - is worse than reporting nothing,
    /// because an agent will act on whichever half suits it.
    /// </summary>
    [SkippableFact]
    public void GetActiveExplorer_SelectionFieldsAgreeWithTheSelectionCount()
    {
        EnsureOutlookAvailable();

        var result = new ApplicationCommands().GetActiveExplorer();

        Assert.True(result.Success, result.ErrorMessage);
        Skip.IfNot(result.HasExplorer, "Outlook is running with no explorer window open.");

        if (result.SelectionCount == 0)
        {
            Assert.Null(result.SelectedItemType);
            Assert.Null(result.SelectedItemMessageClass);
            Assert.Null(result.SelectedItemSubject);
            Assert.False(result.HasMailSelection);
        }
        else
        {
            Assert.False(
                string.IsNullOrWhiteSpace(result.SelectedItemType),
                "Something is selected but its type was not reported.");

            Assert.NotEqual("__ComObject", result.SelectedItemType);
        }

        output.WriteLine($"selection={result.SelectionCount} type={result.SelectedItemType ?? "(none)"}");
    }

    /// <summary>
    /// No inspector is not an error. Asking what is open when nothing is open is a legitimate
    /// question with the answer "nothing", and it must not come back as a failure - an agent that
    /// sees success=false will report the mailbox as broken.
    /// </summary>
    [SkippableFact]
    public void GetActiveInspector_ReportsNothingOpenWithoutFailing()
    {
        EnsureOutlookAvailable();

        var result = new ApplicationCommands().GetActiveInspector();

        Assert.True(result.Success, result.ErrorMessage);

        if (!result.HasInspector)
        {
            Assert.Null(result.ItemType);
            Assert.Null(result.Subject);
            Assert.Null(result.EntryId);
        }

        output.WriteLine($"hasInspector={result.HasInspector} type={result.ItemType ?? "(none)"}");
    }

    /// <summary>
    /// The load-bearing one. Opens a real unsaved compose window and checks the three things the
    /// stranded implementation got wrong:
    ///
    /// 1. the item type must be a real kind, not the runtime-callable wrapper's
    ///    <c>__ComObject</c>, which is what <c>GetType().Name</c> actually returns here;
    /// 2. an unsaved item has no <c>EntryID</c>, and the field must be absent rather than an empty
    ///    string that looks like a handle;
    /// 3. <c>isSaved</c> has to say so, otherwise the caller cannot tell "no id yet" from
    ///    "id withheld".
    ///
    /// The draft is discarded, never sent.
    /// </summary>
    [SkippableFact]
    public void GetActiveInspector_OnAnUnsavedComposeWindow_DoesNotInventAnEntryId()
    {
        EnsureOutlookAvailable();

        string subject = $"{SubjectPrefix}{Guid.NewGuid():N}";

        // Every COM touch below goes through OutlookInteropRunner.Execute so it runs on the shared
        // dispatcher STA thread (ADR-002).
        //
        // Nothing holds the draft's RCW across calls, deliberately. GetActiveInspector
        // final-releases the item it reads, and the CLR keeps one RCW per COM identity per
        // apartment, so a reference held here is the *same* wrapper it disposes of. Holding one
        // made the cleanup path throw "COM object that has been separated from its underlying
        // RCW", which meant the compose window was never closed - and Outlook then autosaved it
        // into Drafts as real mail. Releasing the wrapper does not close the window, so the window
        // is found again by enumerating Inspectors when it is time to close it.
        try
        {
            OutlookInteropRunner.Execute<object?>(
                "test.open-compose-window",
                (application, _) =>
                {
                    OutlookInterop.MailItem? draft = null;

                    try
                    {
                        draft = (OutlookInterop.MailItem)application.CreateItem(OutlookInterop.OlItemType.olMailItem);
                        draft.Subject = subject;
                        draft.Body = "Opened by an integration test. Never sent.";
                        draft.Display(false);
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref draft);
                    }

                    return null;
                },
                ex => throw ex);

            OutlookInspectorContextResult result = WaitForInspector(subject);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(result.HasInspector, "The compose window was displayed but no inspector was reported.");
            Assert.Equal(subject, result.Subject);

            // Probed live: the wrapper's runtime type really is __ComObject, so GetType().Name is
            // not a usable answer to "what kind of item is this".
            Assert.NotEqual("__ComObject", result.ItemType);
            Assert.Equal("IPM.Note", result.MessageClass);

            Assert.False(result.IsSaved, "An unsaved compose window was reported as saved.");

            // Deliberately Assert.Null, not IsNullOrEmpty. The contract is that the field is
            // *absent* - it is serialised with WhenWritingNull, so an empty string would still
            // reach the caller as "entryId": "". An IsNullOrEmpty assertion passes even with the
            // normalisation removed, which sabotage confirmed before this was tightened.
            Assert.Null(result.EntryId);

            output.WriteLine(
                $"type={result.ItemType} messageClass={result.MessageClass} isSaved={result.IsSaved} entryId='{result.EntryId}'");
        }
        finally
        {
            string? cleanupFailure = OutlookInteropRunner.Execute<string?>(
                "test.close-compose-window",
                (application, session) =>
                {
                    CloseComposeWindows(application);
                    SweepLeftoverDrafts(session);
                    return null;
                },
                ex => ex.Message);

            // Rethrowing here would mask the real assertion failure (CA2219). Report it
            // instead, loudly enough that a compose window left on the desktop is not a
            // silent outcome.
            if (cleanupFailure != null)
            {
                output.WriteLine($"WARNING: cleanup did not complete: {cleanupFailure}");
            }
        }
    }

    /// <summary>
    /// Closes any compose window this test class opened, identified by subject rather than by a
    /// held reference. See the note in the test body for why no reference is held.
    /// </summary>
    private static void CloseComposeWindows(OutlookInterop.Application application)
    {
        OutlookInterop.Inspectors? inspectors = null;

        try
        {
            inspectors = application.Inspectors;

            // Backwards: closing one shifts the 1-based collection under a forward walk.
            for (int index = inspectors.Count; index >= 1; index--)
            {
                OutlookInterop.Inspector? inspector = null;
                object? item = null;

                try
                {
                    inspector = inspectors[index];
                    item = inspector.CurrentItem;

                    if (item is OutlookInterop.MailItem mail &&
                        mail.Subject?.StartsWith(SubjectPrefix, StringComparison.Ordinal) == true)
                    {
                        inspector.Close(OutlookInterop.OlInspectorClose.olDiscard);
                    }
                }
                catch (COMException)
                {
                    // A window that cannot be read cannot be identified as ours; leave it alone.
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref item);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref inspectors);
        }
    }

    /// <summary>
    /// Deletes any draft this test class has ever left behind, not just this run's.
    ///
    /// Outlook autosaves an open compose window into Drafts after a short delay, so a run that dies
    /// before its <c>finally</c> - a crashed test host, for instance - leaves real mail in the
    /// user's mailbox. Closing only the item this run created would keep reporting a clean sweep
    /// while those accumulated, which is the failure this project has already hit once.
    /// </summary>
    private static void SweepLeftoverDrafts(OutlookInterop.NameSpace session)
    {
        OutlookInterop.MAPIFolder? drafts = null;
        OutlookInterop.Items? items = null;

        try
        {
            drafts = session.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderDrafts);
            items = drafts.Items;

            // Backwards: deleting shifts the 1-based collection under a forward walk.
            for (int index = items.Count; index >= 1; index--)
            {
                object? entry = null;

                try
                {
                    entry = items[index];
                    if (entry is OutlookInterop.MailItem mail &&
                        mail.Subject?.StartsWith(SubjectPrefix, StringComparison.Ordinal) == true)
                    {
                        mail.Delete();
                    }
                }
                catch (COMException)
                {
                    // An item that cannot be read cannot be one of ours; leave it alone.
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref entry);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref items);
            OutlookInteropRunner.ReleaseComObject(ref drafts);
        }
    }

    private static OutlookInspectorContextResult WaitForInspector(string expectedSubject)
    {
        var commands = new ApplicationCommands();
        OutlookInspectorContextResult result = commands.GetActiveInspector();

        for (int attempt = 0; attempt < 20 && result.Subject != expectedSubject; attempt++)
        {
            Thread.Sleep(250);
            result = commands.GetActiveInspector();
        }

        return result;
    }

    /// <summary>
    /// Probes availability through the dispatcher rather than by resolving the Application here.
    ///
    /// The first version of this helper called <c>TryGetRunningApplication</c> directly and then
    /// released the result with <c>ReleaseComObject</c>, which is <c>FinalReleaseComObject</c>.
    /// That is the shared, already-running Outlook.Application - final-releasing it zeroes the RCW
    /// refcount for *every* holder in the process, so subsequent tests in the same run were left
    /// using a dead wrapper and the test host died with STATUS_STACK_BUFFER_OVERRUN. Exactly the
    /// #19 regression, reintroduced from a test. Do not resolve the Application outside
    /// <see cref="OutlookInteropRunner.Execute"/>.
    /// </summary>
    private static void EnsureOutlookAvailable()
    {
        bool available = OutlookInteropRunner.Execute(
            "test.probe-outlook",
            (_, _) => true,
            _ => false);

        Skip.IfNot(available, "Classic Outlook is not running; start Outlook to exercise this test.");
    }
}
