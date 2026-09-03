using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// The only tests in this repository that send real mail from a real mailbox.
/// <para>
/// Everything else about Outlook automation is recoverable. A send is not: once a message leaves the
/// mailbox it cannot be recalled, and a mistake here reaches a real person. These tests therefore
/// carry a hard constraint that the rest of the suite does not - <b>every message they send is
/// addressed to the signed-in user and to nobody else</b> - and they enforce it in code rather than
/// by convention. <see cref="ResolveOwnSmtpAddress"/> asks Outlook who is signed in, and
/// <see cref="CreateSelfAddressedDraft"/> refuses to save a draft, let alone send one, if the
/// resolved recipient is not that same address.
/// </para>
/// <para>
/// They are marked <c>RunType=OnDemand</c> so no ordinary test run can trigger them by accident.
/// Run them deliberately:
/// <c>dotnet test --filter "Feature=MailSendLifecycle"</c>.
/// </para>
/// <para>
/// This coverage exists because the send, move and delete paths were previously verified only by
/// unit tests of the confirmation gate, which by construction never reach Outlook. Whether a send
/// actually delivers, whether the delivered message can then be found, moved and deleted, and
/// whether the operationId replay cache prevents a duplicate send, were all unproven.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailSendLifecycle")]
[Trait("RequiresOutlook", "true")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public class OutlookSelfSendLifecycleTests(ITestOutputHelper output)
{
    /// <summary>How long to wait for a self-addressed message to come back around via the server.</summary>
    private static readonly TimeSpan DeliveryTimeout = TimeSpan.FromMinutes(3);

    private static readonly TimeSpan DeliveryPollInterval = TimeSpan.FromSeconds(5);

    /// <summary>How long cleanup keeps chasing copies that Exchange has not delivered yet.</summary>
    private static readonly TimeSpan PurgeTimeout = TimeSpan.FromMinutes(3);

    private static readonly TimeSpan PurgePollInterval = TimeSpan.FromSeconds(5);

    /// <summary>
    /// An address that cannot reach anyone. example.com is reserved by RFC 2606 precisely so it can
    /// be used in tests without risk of delivery.
    /// </summary>
    private const string NonDeliverableAddress = "nobody@example.com";

    [SkippableFact]
    public void SendToSelf_ThenFindMoveAndDelete_CompletesTheWholeLifecycle()
    {
        // One test rather than four, because each step needs the message the previous step produced
        // and a half-completed run would leave real mail sitting in a real mailbox. The finally
        // block cleans up whatever stage this got to.
        string ownAddress = ResolveOwnSmtpAddress();
        string token = Guid.NewGuid().ToString("N");

        var commands = new MailCommands();
        var draft = CreateSelfAddressedDraft(commands, ownAddress, token);

        string? deliveredEntryId = null;
        string? deliveredStoreId = null;
        bool sent = false;

        try
        {
            var sendResult = commands.Send(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false,
                confirm: true);

            // Indeterminate is not failure: the Object Model Guard may have put a modal prompt on
            // screen and timed the call out while the message went out anyway. Treat it as sent for
            // cleanup purposes rather than resending, which is exactly the rule the result documents.
            sent = sendResult.Sent || sendResult.Indeterminate;
            Skip.If(
                sendResult.Indeterminate,
                "Send reported an indeterminate outcome, most likely an Object Model Guard prompt. " +
                "A person must dismiss it; the test cannot decide whether the mail went out.");

            Assert.True(sendResult.Success, sendResult.ErrorMessage);
            Assert.True(sendResult.Sent);

            var delivered = WaitForDelivery(commands, token);
            Skip.If(
                delivered is null,
                $"Self-addressed message did not arrive within {DeliveryTimeout.TotalMinutes} minutes. " +
                "That is a mail-flow delay, not a defect in this server.");

            deliveredEntryId = delivered!.EntryId;
            deliveredStoreId = delivered.StoreId;

            Assert.Contains(token, delivered.Subject);

            // Move it out of the Inbox. Drafts is a folder every profile has, so this does not depend
            // on the shape of anyone's folder tree.
            var moveResult = commands.Move(
                targetFolder: "drafts",
                entryId: deliveredEntryId,
                storeId: deliveredStoreId,
                useActiveMail: false);

            Assert.True(moveResult.Success, moveResult.ErrorMessage);

            // Move rewrites the entry ID: the moved item is a different MAPI object in a different
            // store location. Using the stale ID afterwards is a real bug that this asserts against.
            Assert.False(string.IsNullOrWhiteSpace(moveResult.EntryId));

            deliveredEntryId = moveResult.EntryId;
            deliveredStoreId = moveResult.StoreId ?? deliveredStoreId;

            var readBack = commands.Read(
                entryId: deliveredEntryId,
                storeId: deliveredStoreId,
                useActiveMail: false);

            Assert.True(readBack.Success, readBack.ErrorMessage);
            Assert.Contains(token, readBack.Subject);

            var deleteResult = commands.Delete(
                entryId: deliveredEntryId,
                storeId: deliveredStoreId,
                useActiveMail: false);

            Assert.True(deleteResult.Success, deleteResult.ErrorMessage);
            deliveredEntryId = null;

            // Deleted means gone from where it was, not merely reported gone.
            var afterDelete = commands.List(folder: "drafts", maxCount: 50, subjectContains: token);
            Assert.True(afterDelete.Success, afterDelete.ErrorMessage);
            Assert.Empty(afterDelete.Messages);
        }
        finally
        {
            if (!sent && !string.IsNullOrWhiteSpace(draft.EntryId))
            {
                TryDeleteItem(draft.EntryId!, draft.StoreId);
            }

            if (!string.IsNullOrWhiteSpace(deliveredEntryId))
            {
                TryDeleteItem(deliveredEntryId!, deliveredStoreId);
            }

            PurgeStragglers(token, sent);
        }
    }

    /// <summary>
    /// The follow-up flag lifecycle on a genuinely received message (#15).
    ///
    /// <para>
    /// Outlook refuses <c>MarkAsTask</c> and refuses to complete a flag on a draft, so the draft-based
    /// flag tests cannot reach either path. Only a message that has actually been sent and received
    /// can, which is why this lives here with the self-addressing safeguards rather than in the
    /// ordinary suite.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void FlagOnReceivedMail_CanBeRaised_Completed_AndCleared()
    {
        string ownAddress = ResolveOwnSmtpAddress();
        string token = Guid.NewGuid().ToString("N");

        var commands = new MailCommands();
        var draft = CreateSelfAddressedDraft(commands, ownAddress, token);

        string? deliveredEntryId = null;
        string? deliveredStoreId = null;
        bool sent = false;

        try
        {
            var sendResult = commands.Send(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false,
                confirm: true);

            sent = sendResult.Sent || sendResult.Indeterminate;
            Skip.If(
                sendResult.Indeterminate,
                "Send reported an indeterminate outcome, most likely an Object Model Guard prompt.");

            Assert.True(sendResult.Success, sendResult.ErrorMessage);

            var delivered = WaitForDelivery(commands, token);
            Skip.If(
                delivered is null,
                $"Self-addressed message did not arrive within {DeliveryTimeout.TotalMinutes} minutes. " +
                "That is a mail-flow delay, not a defect in this server.");

            deliveredEntryId = delivered!.EntryId;
            deliveredStoreId = delivered.StoreId;

            // A received message starts unflagged, and says so rather than omitting the field.
            Assert.Equal("none", delivered.FlagStatus);

            var due = DateTime.Today.AddDays(3);

            var flagged = commands.SetFlag(
                entryId: deliveredEntryId,
                storeId: deliveredStoreId,
                useActiveMail: false,
                dueDate: due.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                flagRequest: "Review");

            Assert.True(flagged.Success, flagged.ErrorMessage);

            var afterFlag = commands.Read(entryId: deliveredEntryId, storeId: deliveredStoreId, useActiveMail: false);
            Assert.True(afterFlag.Success, afterFlag.ErrorMessage);
            Assert.Equal("flagged", afterFlag.FlagStatus);
            Assert.Equal("Review", afterFlag.FlagRequest);
            Assert.Equal(due, afterFlag.FlagDueDate!.Value.Date);

            // Completing is the path a draft cannot reach at all.
            var completed = commands.SetFlag(
                entryId: deliveredEntryId,
                storeId: deliveredStoreId,
                useActiveMail: false,
                flagStatus: "complete");

            Assert.True(completed.Success, completed.ErrorMessage);

            var afterComplete = commands.Read(entryId: deliveredEntryId, storeId: deliveredStoreId, useActiveMail: false);

            // Completed is not unflagged. Collapsing the two would report handled work as work that
            // was never raised.
            Assert.Equal("complete", afterComplete.FlagStatus);

            var cleared = commands.SetFlag(
                entryId: deliveredEntryId,
                storeId: deliveredStoreId,
                useActiveMail: false,
                flagStatus: "none");

            Assert.True(cleared.Success, cleared.ErrorMessage);

            var afterClear = commands.Read(entryId: deliveredEntryId, storeId: deliveredStoreId, useActiveMail: false);
            Assert.Equal("none", afterClear.FlagStatus);
            Assert.Null(afterClear.FlagDueDate);
        }
        finally
        {
            if (!sent && !string.IsNullOrWhiteSpace(draft.EntryId))
            {
                TryDeleteItem(draft.EntryId!, draft.StoreId);
            }

            if (!string.IsNullOrWhiteSpace(deliveredEntryId))
            {
                TryDeleteItem(deliveredEntryId!, deliveredStoreId);
            }

            PurgeStragglers(token, sent);
        }
    }

    [SkippableFact]
    public void SendToSelf_ReplayedWithSameOperationId_DoesNotSendASecondCopy()
    {
        // #29's replay cache is the difference between a timed-out retry being safe and it duplicating
        // a message. Unit tests only ever exercise it on the refusal path, where nothing is sent and
        // so nothing can be duplicated. This is the first test that proves it on a real send.
        string ownAddress = ResolveOwnSmtpAddress();
        string token = Guid.NewGuid().ToString("N");
        string operationId = Guid.NewGuid().ToString();

        var commands = new MailCommands();
        var draft = CreateSelfAddressedDraft(commands, ownAddress, token);

        bool sent = false;

        try
        {
            var first = commands.Send(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false,
                confirm: true,
                operationId: operationId);

            sent = first.Sent || first.Indeterminate;
            Skip.If(first.Indeterminate, "Send reported an indeterminate outcome; cannot assert replay safely.");
            Assert.True(first.Success, first.ErrorMessage);
            Assert.True(first.Sent);

            // The draft is gone once sent, so a genuine second send could not succeed anyway. What
            // matters is that the replay is answered from the cache and reports success rather than
            // surfacing a "no such item" failure that would tempt a caller into recreating and
            // resending the message.
            //
            // This assertion was confirmed red by giving the replay a fresh operationId: it then
            // fails with "The message you specified cannot be found." The cache is load-bearing.
            var replay = commands.Send(
                entryId: draft.EntryId,
                storeId: draft.StoreId,
                useActiveMail: false,
                confirm: true,
                operationId: operationId);

            Assert.True(replay.Success, replay.ErrorMessage);
            Assert.True(replay.Sent);

            var delivered = WaitForDelivery(commands, token);
            Skip.If(delivered is null, "Self-addressed message did not arrive in time to count copies.");

            // The real assertion: one message, not two.
            var copies = commands.List(folder: "inbox", maxCount: 50, subjectContains: token);
            Assert.True(copies.Success, copies.ErrorMessage);
            Assert.Single(copies.Messages);
        }
        finally
        {
            if (!sent && !string.IsNullOrWhiteSpace(draft.EntryId))
            {
                TryDeleteItem(draft.EntryId!, draft.StoreId);
            }

            PurgeStragglers(token, sent);
        }
    }

    [SkippableFact]
    public void SelfAddressCheck_WhenDraftTargetsSomeoneElse_RefusesBeforeAnythingIsSent()
    {
        // The most important test in this file. Every other test here depends on the guard being
        // real, and a guard that has never been seen to refuse is indistinguishable from one that
        // always passes. This drives it with a recipient that is definitively not the signed-in user
        // and asserts it throws.
        //
        // example.com is reserved by IANA (RFC 2606) and cannot receive mail, so even a catastrophic
        // failure of the guard could not deliver anything to a real person. Nothing here calls Send
        // under any outcome.
        string ownAddress = ResolveOwnSmtpAddress();
        Assert.NotEqual(NonDeliverableAddress, ownAddress);

        string token = Guid.NewGuid().ToString("N");
        var commands = new MailCommands();

        var draft = commands.CreateMailDraft(
            recipientTo: NonDeliverableAddress,
            subject: $"OutlookMcp self-send lifecycle {token}",
            body: "Automated OutlookMcp guard test. Never sent.",
            display: false);

        try
        {
            Assert.True(draft.Success, draft.ErrorMessage);

            var refusal = Assert.Throws<InvalidOperationException>(
                () => AssertDraftIsAddressedOnlyToSelf(draft.EntryId!, draft.StoreId, ownAddress));

            Assert.Contains("Refusing to send", refusal.Message, StringComparison.Ordinal);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(draft.EntryId))
            {
                TryDeleteItem(draft.EntryId!, draft.StoreId);
            }

            // Nothing was sent, by construction, so nothing can be in flight.
            PurgeStragglers(token, messageWasSent: false);
        }
    }

    /// <summary>
    /// Creates a draft addressed to <paramref name="ownAddress"/> and refuses to return it unless
    /// Outlook resolved the recipient to exactly that address.
    /// <para>
    /// The check is on the saved draft rather than on the string passed in, because
    /// <c>Recipients.ResolveAll</c> is what actually decides where the message goes. An unresolved or
    /// differently-resolved recipient is the one failure mode that could send mail to a stranger, so
    /// it throws rather than skips: a skip would look like a benign environment gap.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private MailDraftResult CreateSelfAddressedDraft(MailCommands commands, string ownAddress, string token)
    {
        string subject = $"OutlookMcp self-send lifecycle {token}";

        var draft = commands.CreateMailDraft(
            recipientTo: ownAddress,
            subject: subject,
            body:
                "Automated OutlookMcp integration test. This message was sent by the mailbox owner to " +
                "the mailbox owner and is deleted by the test that created it.",
            display: false);

        Assert.True(draft.Success, draft.ErrorMessage);
        Assert.True(draft.Saved);
        Assert.False(string.IsNullOrWhiteSpace(draft.EntryId));

        AssertDraftIsAddressedOnlyToSelf(draft.EntryId!, draft.StoreId, ownAddress);

        output.WriteLine($"Created self-addressed draft {token} for {ownAddress}.");
        return draft;
    }

    /// <summary>
    /// Reads the saved draft back through COM and throws unless it has exactly one recipient whose
    /// resolved SMTP address is the signed-in user.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void AssertDraftIsAddressedOnlyToSelf(string entryId, string? storeId, string ownAddress)
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            throw new InvalidOperationException("Cannot verify the recipient without a running Outlook.");
        }

        OutlookInterop.NameSpace? session = null;
        object? item = null;
        OutlookInterop.MailItem? mail = null;
        OutlookInterop.Recipients? recipients = null;

        try
        {
            session = application.GetNamespace("MAPI");
            item = session.GetItemFromID(entryId, string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
            mail = item as OutlookInterop.MailItem
                ?? throw new InvalidOperationException("The saved draft is not a mail item.");

            recipients = mail.Recipients;

            if (recipients.Count != 1)
            {
                throw new InvalidOperationException(
                    $"Refusing to send: the draft has {recipients.Count} recipients; exactly one is required.");
            }

            OutlookInterop.Recipient? recipient = null;

            try
            {
                recipient = recipients[1];
                string resolved = ResolveRecipientSmtpAddress(recipient);

                if (!string.Equals(resolved, ownAddress, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Refusing to send: the draft's recipient did not resolve to the signed-in user.");
                }
            }
            finally
            {
                ReleaseComObject(recipient);
            }
        }
        finally
        {
            ReleaseComObject(recipients);
            ReleaseComObject(mail);
            ReleaseComObject(item);
            ReleaseComObject(session);
            ReleaseSharedApplication(application);
        }
    }

    /// <summary>
    /// Asks Outlook for the signed-in user's primary SMTP address. Exchange accounts expose an X.500
    /// address on <c>AddressEntry.Address</c>, so the Exchange user object is preferred and the raw
    /// address is only accepted when it already looks like SMTP.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private string ResolveOwnSmtpAddress()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping self-send test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        OutlookInterop.NameSpace? session = null;
        OutlookInterop.Recipient? currentUser = null;
        OutlookInterop.AddressEntry? entry = null;
        OutlookInterop.ExchangeUser? exchangeUser = null;

        try
        {
            session = application.GetNamespace("MAPI");
            currentUser = session.CurrentUser;
            entry = currentUser?.AddressEntry;

            string? address = null;

            exchangeUser = entry?.GetExchangeUser();
            if (exchangeUser is not null)
            {
                address = exchangeUser.PrimarySmtpAddress;
            }

            if (string.IsNullOrWhiteSpace(address) && entry?.Address is string raw && raw.Contains('@', StringComparison.Ordinal))
            {
                address = raw;
            }

            if (string.IsNullOrWhiteSpace(address) || !address.Contains('@', StringComparison.Ordinal))
            {
                throw new SkipException(
                    "Could not resolve the signed-in user's SMTP address; refusing to send anything.");
            }

            return address;
        }
        catch (SkipException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new SkipException($"Outlook MAPI session is not usable: {ex.Message}", ex);
        }
        finally
        {
            ReleaseComObject(exchangeUser);
            ReleaseComObject(entry);
            ReleaseComObject(currentUser);
            ReleaseComObject(session);
            ReleaseSharedApplication(application);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string ResolveRecipientSmtpAddress(OutlookInterop.Recipient recipient)
    {
        OutlookInterop.AddressEntry? entry = null;
        OutlookInterop.ExchangeUser? exchangeUser = null;

        try
        {
            if (!recipient.Resolved)
            {
                recipient.Resolve();
            }

            if (!recipient.Resolved)
            {
                throw new InvalidOperationException("Refusing to send: the recipient did not resolve.");
            }

            entry = recipient.AddressEntry;
            exchangeUser = entry?.GetExchangeUser();

            if (exchangeUser?.PrimarySmtpAddress is string smtp && !string.IsNullOrWhiteSpace(smtp))
            {
                return smtp;
            }

            if (entry?.Address is string raw && raw.Contains('@', StringComparison.Ordinal))
            {
                return raw;
            }

            throw new InvalidOperationException(
                "Refusing to send: the recipient's SMTP address could not be determined.");
        }
        finally
        {
            ReleaseComObject(exchangeUser);
            ReleaseComObject(entry);
        }
    }

    /// <summary>
    /// Polls the Inbox for the self-addressed message, using the server-side subject filter so the
    /// poll stays cheap and finds the message regardless of how much other mail arrives meanwhile.
    /// </summary>
    private MailSummaryInfo? WaitForDelivery(MailCommands commands, string token)
    {
        DateTimeOffset deadline = DateTimeOffset.UtcNow + DeliveryTimeout;

        while (DateTimeOffset.UtcNow < deadline)
        {
            var result = commands.List(folder: "inbox", maxCount: 25, subjectContains: token);
            Assert.True(result.Success, result.ErrorMessage);

            if (result.Messages.Count > 0)
            {
                output.WriteLine($"Message {token} arrived after polling.");
                return result.Messages[0];
            }

            Thread.Sleep(DeliveryPollInterval);
        }

        output.WriteLine($"Message {token} never arrived.");
        return null;
    }

    /// <summary>
    /// Last-resort cleanup. Sweeps the folders a test message can end up in and removes anything
    /// still carrying this run's token, so a failure part-way through does not leave real mail behind.
    /// <para>
    /// Deleted Items is swept last and deliberately: <c>MailItem.Delete</c> is a soft delete, so a
    /// message the test already "deleted" is sitting there, and a self-addressed send leaves two
    /// copies (the sent one and the received one). Deleting an item that is already in Deleted Items
    /// removes it for good, which leaves the mailbox as the test found it.
    /// </para>
    /// <para>
    /// A single sweep is not enough and this was observed, not theorised: a self-addressed send
    /// round-trips through Exchange, so the Sent and Inbox copies can materialise seconds after the
    /// test body has finished. An immediate sweep finds nothing, returns happy, and leaves real mail
    /// sitting in a real corporate mailbox. When <paramref name="messageWasSent"/> is true this
    /// therefore keeps sweeping until two consecutive passes come back clean, or until it runs out of
    /// patience. When nothing was sent there is nothing in flight and one pass is enough.
    /// </para>
    /// </summary>
    private void PurgeStragglers(string token, bool messageWasSent)
    {
        if (!messageWasSent)
        {
            SweepOnce(token);
            return;
        }

        DateTimeOffset deadline = DateTimeOffset.UtcNow.Add(PurgeTimeout);
        int consecutiveCleanSweeps = 0;

        while (DateTimeOffset.UtcNow < deadline && consecutiveCleanSweeps < 2)
        {
            consecutiveCleanSweeps = SweepOnce(token) == 0 ? consecutiveCleanSweeps + 1 : 0;

            if (consecutiveCleanSweeps < 2)
            {
                Thread.Sleep(PurgePollInterval);
            }
        }

        if (consecutiveCleanSweeps < 2)
        {
            output.WriteLine(
                $"WARNING: gave up purging {token} after {PurgeTimeout.TotalSeconds:F0}s. " +
                "Check Inbox, Sent Items and Deleted Items for leftovers.");
        }
    }

    /// <summary>
    /// One pass over every folder a test message can land in. Returns how many items it removed, so
    /// the caller can tell a quiet mailbox from one that is still receiving copies.
    /// </summary>
    private int SweepOnce(string token)
    {
        var commands = new MailCommands();
        int purged = 0;

        foreach (string folder in new[] { "inbox", "drafts", "sent", "outbox", "deleted" })
        {
            try
            {
                var found = commands.List(folder: folder, maxCount: 50, subjectContains: token);

                if (!found.Success)
                {
                    continue;
                }

                foreach (var message in found.Messages)
                {
                    if (!string.IsNullOrWhiteSpace(message.EntryId))
                    {
                        TryDeleteItem(message.EntryId!, message.StoreId);
                        purged++;
                        output.WriteLine($"Purged a straggler from {folder}.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Cleanup must never mask the assertion that actually failed.
                output.WriteLine($"Could not purge {folder}: {ex.Message}");
            }
        }

        return purged;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void TryDeleteItem(string entryId, string? storeId)
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            return;
        }

        OutlookInterop.NameSpace? session = null;
        object? item = null;
        OutlookInterop.MailItem? mail = null;

        try
        {
            session = application.GetNamespace("MAPI");
            item = session.GetItemFromID(entryId, string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
            mail = item as OutlookInterop.MailItem;
            mail?.Delete();
        }
        catch (Exception)
        {
            // The item may already be gone, which is the outcome cleanup wanted anyway.
        }
        finally
        {
            ReleaseComObject(mail);
            ReleaseComObject(item);
            ReleaseComObject(session);
            ReleaseSharedApplication(application);
        }
    }

    /// <summary>
    /// The Application obtained via <c>TryGetRunningApplication</c> is the user's shared,
    /// already-running Outlook instance. Decrement its ref count; never final-release it, or the
    /// cached RCW is invalidated for every other holder in the process. See #19.
    /// </summary>
    private static void ReleaseSharedApplication(object? value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.ReleaseComObject(value);
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }
    }
}
