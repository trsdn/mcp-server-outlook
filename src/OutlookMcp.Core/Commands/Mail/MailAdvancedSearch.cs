using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookMcp.ComInterop;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// Runs <c>Application.AdvancedSearch</c> and hands back its result rowset (#13, #27).
///
/// <para>
/// <b>Why this engine exists.</b> <c>Items.Restrict</c> cannot filter on <c>Body</c>, <c>HTMLBody</c>,
/// <c>EntryID</c>, <c>RecurrenceState</c>, <c>Saved</c> or <c>Sent</c> - MAPI simply will not accept a
/// restriction over them. A free-text search therefore had to open every candidate message and check
/// the body in this process, and give up at a scan limit. In a large folder that limit is not a
/// performance detail: a match past it is reported to the caller as "no such mail".
/// <c>AdvancedSearch</c> asks Outlook to run the same substring question itself, over the whole
/// folder, with no client-side scan and no horizon.
/// </para>
///
/// <para>
/// <b>The completion problem, and how it is solved.</b> <c>AdvancedSearch</c> is the only
/// asynchronous call this server makes. It returns a <c>Search</c> object immediately, before the
/// results exist, and signals completion by raising <c>AdvancedSearchComplete</c> on the apartment
/// that registered the handler. That apartment is the process-wide STA dispatcher thread (ADR-002),
/// and out-of-process COM events reach an STA as window messages. The dispatcher thread does not
/// normally pump: it blocks on its work-item channel. So a work item that simply waits for the event
/// would block the very thread the event has to be delivered on, and would sit there until the
/// operation timeout expired.
/// </para>
///
/// <para>
/// The fix is to pump the STA message queue while waiting, via <see cref="StaMessagePump"/>. Three
/// properties make that safe here rather than merely expedient:
/// </para>
/// <list type="number">
/// <item><description>
/// <b>Pumping cannot re-enter another Outlook operation.</b> The dispatcher's queue is a
/// <c>Channel</c>, not a window message queue. Draining window messages delivers Outlook's callbacks
/// and nothing else; the next queued work item stays queued behind the current one exactly as before.
/// </description></item>
/// <item><description>
/// <b>The handler is matched on <c>Tag</c>.</b> The event carries whichever search completed, and a
/// user-initiated search in the Outlook UI raises it too. A handler that assumed the event was its
/// own would return another search's results.
/// </description></item>
/// <item><description>
/// <b>The handler is unhooked in a <c>finally</c>.</b> A leaked sink on the shared, long-lived
/// <c>Application</c> would keep firing into a dead closure for the life of the process.
/// </description></item>
/// </list>
///
/// <para>
/// <b>When it does not finish.</b> The wait is bounded well inside the dispatcher's own timeout, so a
/// slow search degrades into a described partial answer rather than into a stalled dispatcher. The
/// search is stopped, whatever it found is returned, and the caller is told the search was
/// incomplete - never handed a short result set that looks exhaustive.
/// </para>
/// </summary>
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
internal static class MailAdvancedSearch
{
    /// <summary>Value of <c>searchEngine</c> for results this engine produced.</summary>
    public const string EngineName = "advancedSearch";

    /// <summary>
    /// How long to wait for <c>AdvancedSearchComplete</c>.
    ///
    /// <para>
    /// Deliberately far below <see cref="ComInteropConstants.DefaultOperationTimeout"/>. Exceeding
    /// the dispatcher timeout instead would abandon the caller's wait while the search kept running
    /// on the STA thread, wedging every later operation behind it (see ADR-002, decision 4). A search
    /// over a 2,400-item folder completed in under three seconds when this was measured, so a minute
    /// is a wide margin rather than a tight one.
    /// </para>
    /// </summary>
    private static readonly TimeSpan CompletionTimeout = TimeSpan.FromSeconds(60);

    /// <summary>
    /// The outcome of one search: the result table, and whether the search actually finished.
    /// </summary>
    /// <param name="Table">The result rowset. The caller owns it and must release it.</param>
    /// <param name="Completed">
    /// False when the search was stopped at the timeout. The rows are real but the set is partial,
    /// which the caller must report rather than absorb.
    /// </param>
    public sealed record Result(Outlook.Table Table, bool Completed);

    /// <summary>
    /// Runs a search over one folder and returns its rowset, or <see langword="null"/> if this search
    /// could not be started at all.
    ///
    /// <para>
    /// A null return is not a failure of the request - it means this engine cannot answer it, and the
    /// caller falls back to the client-side scan and says so in the response.
    /// </para>
    /// </summary>
    /// <param name="application">The shared Outlook application. Never released here; it is not ours.</param>
    /// <param name="folder">Folder to search. Subfolders are excluded, matching folder-scoped listing semantics.</param>
    /// <param name="filter">A DASL filter with no <c>@SQL=</c> prefix.</param>
    /// <param name="configureTable">Applies the projected columns and sort order to the result table.</param>
    /// <param name="unavailableReason">Set when this returns null, describing why for the response.</param>
    public static Result? TryRun(
        Outlook.Application application,
        Outlook.MAPIFolder folder,
        string filter,
        Action<Outlook.Table> configureTable,
        out string? unavailableReason)
    {
        unavailableReason = null;

        string? scope = BuildScope(folder);
        if (scope == null)
        {
            unavailableReason =
                "this folder's path contains a quote character, which cannot be expressed unambiguously in an "
                + "AdvancedSearch scope";
            return null;
        }

        // Identifies our own search in the completion event. Outlook raises AdvancedSearchComplete
        // for every search in the process, including ones a person started in the Outlook window, so
        // an untagged handler would happily accept somebody else's results.
        string tag = "OutlookMcp-" + Guid.NewGuid().ToString("N");
        bool completed = false;

        void OnComplete(Outlook.Search search)
        {
            if (string.Equals(SafeTag(search), tag, StringComparison.Ordinal))
            {
                completed = true;
            }
        }

        var handler = new Outlook.ApplicationEvents_11_AdvancedSearchCompleteEventHandler(OnComplete);
        Outlook.Search? running = null;
        Outlook.Table? table = null;
        bool handing = false;

        try
        {
            application.AdvancedSearchComplete += handler;

            try
            {
                running = application.AdvancedSearch(scope, filter, false, tag);
            }
            catch (COMException ex)
            {
                // Outlook rejected the scope or the filter. The request is still perfectly
                // answerable by the scan, so this is a fallback rather than a failure. Narrow and
                // deliberate: see Rule 1b's allowance for a genuinely different code path.
                unavailableReason = ex.Message;
                return null;
            }

            bool finished = StaMessagePump.WaitFor(() => completed, CompletionTimeout);

            if (!finished)
            {
                // Stop it rather than leaving it running: an abandoned search keeps consuming the
                // store's attention, and its later completion event would arrive with nothing left
                // to receive it.
                TryStop(running);
            }

            try
            {
                table = running.GetTable();
                configureTable(table);
            }
            catch (COMException ex)
            {
                // The search itself ran, but this store will not hand back a result rowset, will not
                // accept one of the projected columns, or will not sort it. That is the same
                // condition the folder listing treats as "this store cannot answer from a rowset",
                // and it must be treated the same way here: the request is still perfectly
                // answerable by the client-side scan, so fall back rather than failing a search that
                // Outlook was willing to run. Letting this propagate would turn every default
                // mail.search on such a store into an error.
                unavailableReason = ex.Message;
                return null;
            }

            var result = new Result(table, finished);
            handing = true;
            return result;
        }
        finally
        {
            // A sink left attached to the shared, process-lifetime Application would go on firing
            // into a dead closure forever.
            try
            {
                application.AdvancedSearchComplete -= handler;
            }
            catch (COMException)
            {
                // Detaching can fail if Outlook is tearing the connection point down. There is
                // nothing to recover and nothing the caller can do about it, and the search itself
                // already succeeded or failed on its own terms.
            }

            if (!handing)
            {
                OutlookInterop.OutlookInteropRunner.ReleaseComObject(ref table);
            }

            OutlookInterop.OutlookInteropRunner.ReleaseComObject(ref running);
        }
    }

    /// <summary>
    /// The message put on a partial result. Written once, here, so the two projections cannot drift
    /// into describing the same incompleteness differently.
    /// </summary>
    public static string DescribeIncompleteSearch()
        => string.Format(
            CultureInfo.InvariantCulture,
            "Outlook did not finish this search within {0} seconds, so it was stopped and these are the matches "
            + "found so far, not the whole result set. Do not read this as 'no further matches exist'. Narrow it "
            + "with structured filters (folder, fromAddress, a date range), or use searchMode 'fullText' to have "
            + "the content index answer instead.",
            (int)CompletionTimeout.TotalSeconds);

    /// <summary>
    /// The message put on a response that fell back to the client-side scan.
    /// </summary>
    public static string DescribeFallback(string reason)
        => string.Format(
            CultureInfo.InvariantCulture,
            "This query could not be handed to Outlook's own search engine, so it was run as a client-side scan "
            + "instead: each candidate message was opened and checked, which is slower and stops at a scan limit, "
            + "so a match far back in a very large folder may be missed. Reason: {0}",
            reason);

    /// <summary>
    /// Builds the scope argument: a single-quoted folder path.
    ///
    /// <para>
    /// Returns <see langword="null"/> for a path containing a quote. Outlook documents no escape for
    /// the scope string, so a folder named <c>Anna's</c> would produce a scope that either fails or,
    /// worse, parses as something else. Declining is the honest option; the caller falls back.
    /// </para>
    /// </summary>
    private static string? BuildScope(Outlook.MAPIFolder folder)
    {
        string? path;

        try
        {
            path = folder.FolderPath;
        }
        catch (COMException)
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(path) || path.Contains('\'', StringComparison.Ordinal))
        {
            return null;
        }

        return "'" + path + "'";
    }

    private static string? SafeTag(Outlook.Search search)
    {
        try
        {
            return search.Tag;
        }
        catch (COMException)
        {
            // A search raised by something else in the process that we cannot inspect. Treating it
            // as "not ours" is the only safe reading.
            return null;
        }
    }

    private static void TryStop(Outlook.Search search)
    {
        try
        {
            search.Stop();
        }
        catch (COMException)
        {
            // Already finished, or Outlook will not stop it. Either way the results below are what
            // there is, and the caller is told the search was incomplete regardless.
        }
    }
}
