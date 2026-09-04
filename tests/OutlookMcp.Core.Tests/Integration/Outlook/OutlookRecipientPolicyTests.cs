using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// The opt-in recipient allow-list, exercised against a real draft (#9).
///
/// <para>
/// <c>RecipientPolicyTests</c> covers the matching rules as pure logic. What it cannot cover is the
/// half that only exists against Outlook: reading the recipients off a real <c>MailItem</c>, where
/// an internal Exchange recipient's <c>Address</c> is an X500 path rather than an SMTP address and
/// has to be resolved through the address book.
/// </para>
///
/// <para>
/// <b>Nothing is ever sent here.</b> Every test asserts a <i>refusal</i>, then verifies the draft is
/// still unsent. Confirming the allowed path would mean putting real mail in a real Outbox from a
/// mailbox shared with other work, which is not worth the coverage.
/// </para>
///
/// <para>
/// <b>Note for anyone re-running these against a broken build.</b> A test that asserts "this must
/// be refused" does send the message when the refusal is missing - which is exactly what happened
/// on the red run that drove this feature. The recipient is deliberately in the reserved
/// <c>.invalid</c> TLD (RFC 2606) so nothing can ever be delivered to a real person, but the
/// message does reach Sent Items and has to be cleaned up by hand.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "RecipientPolicy")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookRecipientPolicyTests(ITestOutputHelper output) : IDisposable
{
    private const string Marker = "OutlookMcp recipient policy test";

    private readonly string? _originalPolicy =
        Environment.GetEnvironmentVariable(RecipientPolicy.EnvironmentVariableName);

    public void Dispose()
    {
        Environment.SetEnvironmentVariable(RecipientPolicy.EnvironmentVariableName, _originalPolicy);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// A confirmed send to an address outside the policy must be refused before
    /// <c>MailItem.Send()</c> is reached, and the draft must survive as a draft.
    /// </summary>
    [SkippableFact]
    public void Send_ToARecipientOutsideThePolicy_IsRefusedAndTheDraftSurvives()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string subject = $"{Marker} {Guid.NewGuid():N}";
        string? draftId = null;

        Environment.SetEnvironmentVariable(
            RecipientPolicy.EnvironmentVariableName, "@allowed-by-nobody.invalid");

        try
        {
            var draft = commands.CreateMailDraft(
                recipientTo: "blocked@definitely-not-allowed.invalid",
                subject: subject,
                body: "This draft must never be sent.");

            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;
            Assert.False(string.IsNullOrWhiteSpace(draftId));

            var refused = commands.Send(entryId: draftId, useActiveMail: false, confirm: true);

            output.WriteLine($"Refused as expected: {refused.ErrorMessage}");

            Assert.False(refused.Success);
            Assert.False(refused.Sent);
            Assert.False(refused.Indeterminate);
            Assert.NotNull(refused.ErrorMessage);
            Assert.Contains(
                RecipientPolicy.EnvironmentVariableName,
                refused.ErrorMessage!,
                StringComparison.Ordinal);

            // The refusal has to be a refusal. A send that reported one and went anyway would pass
            // every assertion above.
            var readBack = commands.Read(entryId: draftId, useActiveMail: false);
            Assert.True(readBack.Success, readBack.ErrorMessage);
            Assert.Equal(subject, readBack.Subject);
        }
        finally
        {
            if (draftId != null)
            {
                _ = commands.Delete(entryId: draftId, useActiveMail: false);
            }
        }
    }

    /// <summary>
    /// With no policy configured, send must reach Outlook exactly as it did before this feature
    /// existed. Proven without sending anything: the draft has no recipients at all, so Outlook
    /// itself refuses it - and the message that comes back must be Outlook's, not the policy's.
    /// </summary>
    [SkippableFact]
    public void Send_WithNoPolicyConfigured_IsNotIntercepted()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string subject = $"{Marker} {Guid.NewGuid():N}";
        string? draftId = null;

        Environment.SetEnvironmentVariable(RecipientPolicy.EnvironmentVariableName, null);

        try
        {
            var draft = commands.CreateMailDraft(subject: subject, body: "No recipients on purpose.");
            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            var result = commands.Send(entryId: draftId, useActiveMail: false, confirm: true);

            output.WriteLine($"Result: success={result.Success} sent={result.Sent} {result.ErrorMessage}");

            Assert.False(result.Success);
            Assert.False(result.Sent);
            Assert.NotNull(result.ErrorMessage);
            Assert.DoesNotContain(
                RecipientPolicy.EnvironmentVariableName,
                result.ErrorMessage!,
                StringComparison.Ordinal);
        }
        finally
        {
            if (draftId != null)
            {
                _ = commands.Delete(entryId: draftId, useActiveMail: false);
            }
        }
    }

    /// <summary>
    /// A refusal must not be written into the send idempotency cache.
    ///
    /// <para>
    /// The cache exists so a retry cannot duplicate a message the first attempt may already have
    /// sent (#29). A policy refusal sent nothing, and its own error text invites a retry once the
    /// recipients or the allow-list are fixed - so replaying it would answer that retry from a
    /// stale result and make the fix look ineffective.
    /// </para>
    ///
    /// <para>
    /// Proven without ever sending: the second attempt is made under a <i>different</i> policy that
    /// still refuses. If the result were replayed the message would name the first policy.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void Send_RefusedByPolicy_IsNotCachedAgainstItsOperationId()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();
        string subject = $"{Marker} {Guid.NewGuid():N}";
        string operationId = Guid.NewGuid().ToString();
        string? draftId = null;

        try
        {
            var draft = commands.CreateMailDraft(
                recipientTo: "blocked@definitely-not-allowed.invalid",
                subject: subject,
                body: "This draft must never be sent.");

            Assert.True(draft.Success, draft.ErrorMessage);
            draftId = draft.EntryId;

            Environment.SetEnvironmentVariable(
                RecipientPolicy.EnvironmentVariableName, "@first-policy.invalid");

            var first = commands.Send(
                entryId: draftId, useActiveMail: false, confirm: true, operationId: operationId);

            Assert.False(first.Success);
            Assert.False(first.Sent);
            Assert.Contains("@first-policy.invalid", first.ErrorMessage!, StringComparison.Ordinal);

            // Same operationId, different policy - and still one that refuses, so nothing is sent.
            Environment.SetEnvironmentVariable(
                RecipientPolicy.EnvironmentVariableName, "@second-policy.invalid");

            var second = commands.Send(
                entryId: draftId, useActiveMail: false, confirm: true, operationId: operationId);

            output.WriteLine($"Second attempt: {second.ErrorMessage}");

            Assert.False(second.Success);
            Assert.False(second.Sent);
            Assert.Contains("@second-policy.invalid", second.ErrorMessage!, StringComparison.Ordinal);
            Assert.DoesNotContain("@first-policy.invalid", second.ErrorMessage!, StringComparison.Ordinal);
        }
        finally
        {
            if (draftId != null)
            {
                _ = commands.Delete(entryId: draftId, useActiveMail: false);
            }
        }
    }

    /// <summary>
    /// The allow-list must cover meeting invitations too. <c>calendar create-appointment</c> with
    /// <c>sendInvitation</c> is the second path that puts a message addressed to caller-chosen
    /// recipients outside the mailbox; guarding only <c>mail.send</c> would leave the user believing
    /// nothing could reach an unlisted address while a second route stayed open.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_InvitingSomeoneOutsideThePolicy_IsRefused()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string subject = $"{Marker} {Guid.NewGuid():N}";
        DateTimeOffset start = DateTimeOffset.Now.Date.AddDays(30).AddHours(9);

        Environment.SetEnvironmentVariable(
            RecipientPolicy.EnvironmentVariableName, "@allowed-by-nobody.invalid");

        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = commands.CreateAppointment(
                subject: subject,
                start: start.ToString("o"),
                endTime: start.AddMinutes(30).ToString("o"),
                requiredAttendees: "outsider@definitely-not-allowed.invalid",
                sendInvitation: true);

            output.WriteLine($"Refused as expected: {created.ErrorMessage}");

            entryId = created.EntryId;
            storeId = created.StoreId;

            Assert.False(created.Success);
            Assert.False(created.InvitationSent);
            Assert.NotNull(created.ErrorMessage);
            Assert.Contains(
                RecipientPolicy.EnvironmentVariableName,
                created.ErrorMessage!,
                StringComparison.Ordinal);

            // The appointment itself is saved and the caller is told so, precisely so they do not
            // create it a second time believing nothing happened.
            Assert.True(created.Saved);
        }
        finally
        {
            if (entryId != null)
            {
                _ = commands.DeleteAppointment(
                    entryId: entryId, storeId: storeId, useActiveAppointment: false);
            }
        }
    }

    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
    }
}
