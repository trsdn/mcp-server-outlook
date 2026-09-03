using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Responding to a meeting invitation (#32).
///
/// <para>
/// <b>These tests never accept anything.</b> Accepting a genuine invitation would put a real event in
/// the owner's calendar and, if a response were sent, mail a real organiser. There is no way to
/// manufacture a safe invitation - you cannot invite yourself and then answer it - so what is covered
/// here is every path that refuses: no item, the wrong kind of item, an unknown response value.
/// </para>
///
/// <para>
/// The accepting path itself is exercised by <c>RespondToMeeting_AcceptsTheNominatedInvitation</c>,
/// which runs only when <c>OUTLOOKMCP_RESPOND_ENTRYID</c> names an invitation the operator has
/// deliberately chosen. Until somebody runs that, accept/decline/tentative is implemented and
/// unverified, and the PR and issue say so rather than implying otherwise.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailMeetingRespond")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMeetingRespondTests(ITestOutputHelper output)
{
    private static readonly string[] FolderCandidates =
        ["inbox", "Inbox/older", "Inbox/yesterbox", "Archive"];

    /// <summary>
    /// Nothing to respond to must be an error, not a quiet no-op that reads as success.
    /// </summary>
    [SkippableFact]
    public void RespondToMeeting_WithoutAnItem_Fails()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().RespondToMeeting(response: "accept");

        Assert.False(result.Success);
        Assert.False(result.ResponseSent);
        Assert.NotNull(result.ErrorMessage);
    }

    /// <summary>
    /// An unknown response value must be rejected before anything is resolved. Guessing between
    /// accept and decline on the caller's behalf is not a recoverable mistake.
    /// </summary>
    [SkippableFact]
    public void RespondToMeeting_WithUnknownResponse_Fails()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().RespondToMeeting(entryId: "irrelevant", response: "maybe-ish");

        Assert.False(result.Success);
        Assert.False(result.ResponseSent);
        Assert.Contains("maybe-ish", result.ErrorMessage!, StringComparison.Ordinal);
    }

    /// <summary>
    /// Pointing it at an ordinary message must say so plainly. The item is only read; nothing about
    /// it is changed.
    /// </summary>
    [SkippableFact]
    public void RespondToMeeting_OnAnOrdinaryMessage_FailsAndSaysWhy()
    {
        var (entryId, storeId) = FindItemOfType("mail");

        var result = new MailCommands().RespondToMeeting(
            entryId: entryId,
            storeId: storeId,
            response: "accept");

        Assert.False(result.Success);
        Assert.False(result.ResponseSent);
        Assert.NotNull(result.ErrorMessage);

        output.WriteLine($"Rejected as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// A meeting <em>response</em> is not a meeting <em>request</em>, and neither is a cancellation.
    /// Both look like scheduling items in a listing, so an agent will reach for them.
    /// </summary>
    [SkippableFact]
    public void RespondToMeeting_OnAMeetingResponse_FailsAndSaysWhy()
    {
        var (entryId, storeId) = FindItemOfType("meetingResponse");

        var result = new MailCommands().RespondToMeeting(
            entryId: entryId,
            storeId: storeId,
            response: "decline");

        Assert.False(result.Success);
        Assert.False(result.ResponseSent);
        Assert.NotNull(result.ErrorMessage);

        output.WriteLine($"Rejected as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// The real path, run only against an invitation the operator nominated by entry id. It accepts
    /// without sending a response, so no organiser is mailed; the calendar entry it creates is the
    /// operator's to keep or remove.
    /// </summary>
    [SkippableFact]
    [Trait("RunType", "OnDemand")]
    public void RespondToMeeting_AcceptsTheNominatedInvitation()
    {
        string? entryId = Environment.GetEnvironmentVariable("OUTLOOKMCP_RESPOND_ENTRYID");

        Skip.If(
            string.IsNullOrWhiteSpace(entryId),
            "Set OUTLOOKMCP_RESPOND_ENTRYID to an invitation you are willing to accept. "
                + "This test is not run by default because accepting a real invitation is not reversible.");

        var result = new MailCommands().RespondToMeeting(
            entryId: entryId,
            response: "tentative",
            sendResponse: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.ResponseSent);
        Assert.Equal("tentative", result.Response);

        output.WriteLine($"Responded tentatively to '{result.Subject}' without notifying the organiser.");
    }

    /// <summary>
    /// Finds a real item of the given <c>itemType</c>, skipping when the mailbox has none. A test
    /// that cannot find its subject has verified nothing and must say so rather than pass.
    /// </summary>
    private static (string EntryId, string? StoreId) FindItemOfType(string itemType)
    {
        var commands = new MailCommands();

        foreach (string folder in FolderCandidates)
        {
            var listed = commands.List(folder: folder, maxCount: 100);

            if (!listed.Success)
            {
                continue;
            }

            var match = listed.Messages.FirstOrDefault(
                m => m.ItemType == itemType && !string.IsNullOrWhiteSpace(m.EntryId));

            if (match is not null)
            {
                return (match.EntryId!, match.StoreId);
            }
        }

        throw new SkipException($"This mailbox holds no item of type '{itemType}' to point the test at.");
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping meeting respond test: no running classic Outlook desktop instance is available.");
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
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            if (session != null && Marshal.IsComObject(session))
            {
                _ = Marshal.FinalReleaseComObject(session);
            }
        }
    }
}
