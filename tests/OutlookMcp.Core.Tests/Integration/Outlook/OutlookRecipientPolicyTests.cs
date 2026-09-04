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
