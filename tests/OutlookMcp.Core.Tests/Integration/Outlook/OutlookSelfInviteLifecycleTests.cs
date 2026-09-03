using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// The one path in the attendee feature that actually mails somebody: <c>sendInvitation</c> (#32).
///
/// <para>
/// This sends a real meeting invitation through the signed-in account. The only attendee is the
/// mailbox owner, and that is enforced in code rather than by convention: the attendee list is read
/// back after Outlook resolves it and the test refuses to send unless every resolved attendee is the
/// owner. Marked <c>RunType=OnDemand</c> so no ordinary run can trigger it by accident.
/// </para>
///
/// <para>
/// <b>What this does not prove.</b> With the owner as the only attendee there is nobody for Exchange
/// to notify, and this mailbox keeps no local Sent Items, so delivery was checked for and not found.
/// The test therefore establishes that the send path executes against real Outlook and that the
/// self-addressing guard holds - not that an invitation reached an inbox. Saying otherwise would be
/// the same "green without having verified anything" this project keeps tripping over.
/// </para>
///
/// <para>
/// It exists because the alternative is shipping a send path that has never sent anything. The
/// sibling <see cref="OutlookMeetingAttendeeTests"/> covers everything that does not leave the
/// machine.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CalendarAttendees")]
[Trait("RequiresOutlook", "true")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public class OutlookSelfInviteLifecycleTests(ITestOutputHelper output)
{
    /// <summary>
    /// Sends a meeting invitation addressed only to the mailbox owner, then removes the meeting.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_SendInvitationToSelf_Sends()
    {
        OwnerAddresses owner = ResolveOwnAddresses();
        string own = owner.Smtp;
        string subject = $"OutlookMcp self-invite {Guid.NewGuid():N}";
        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(9).Date.AddHours(10);

        // Resolve the attendee list before sending anything: if Outlook resolves the address to
        // anybody other than the owner, this must not send.
        var rehearsal = commands.CreateAppointment(
            subject: $"{subject} (rehearsal)",
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            requiredAttendees: own);

        try
        {
            Assert.True(rehearsal.Success, rehearsal.ErrorMessage);
            AssertOnlyAddressedToSelf(rehearsal.Attendees.Select(a => a.Address), owner);
        }
        finally
        {
            DeleteCreatedItem(commands, rehearsal.EntryId, rehearsal.StoreId);
        }

        var sent = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            body: "Sent by the OutlookMcp self-invite test. Safe to delete.",
            requiredAttendees: own,
            sendInvitation: true);

        try
        {
            Assert.True(sent.Success, sent.ErrorMessage);
            Assert.True(sent.IsMeeting);
            Assert.True(sent.InvitationSent);
            Assert.Empty(sent.UnresolvedAttendees);

            output.WriteLine($"Sent self-invitation '{subject}'.");
        }
        finally
        {
            DeleteCreatedItem(commands, sent.EntryId, sent.StoreId);
        }
    }

    /// <summary>
    /// Refusing to invite nobody. <c>sendInvitation</c> without attendees is a caller mistake, and
    /// silently creating a private appointment instead would hide it.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_SendInvitationWithoutAttendees_Fails()
    {
        EnsureOutlookAvailable();

        DateTimeOffset start = DateTimeOffset.Now.AddDays(9).Date.AddHours(16);

        var result = new CalendarCommands().CreateAppointment(
            subject: $"OutlookMcp no-attendee invite {Guid.NewGuid():N}",
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            sendInvitation: true);

        Assert.False(result.Success);
        Assert.False(result.Saved);
        Assert.False(result.InvitationSent);
        Assert.Contains("no attendees", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Throws unless every resolved attendee is the mailbox owner. The constraint lives in code
    /// because "the test only ever invites me" is exactly the kind of promise that quietly stops
    /// being true.
    /// </summary>
    /// <remarks>
    /// Exchange resolves a recipient to an X.500 address, not SMTP, so both forms of the owner's own
    /// address are compared - exactly, never by substring. An earlier version matched the SMTP local
    /// part against the X.500 <c>cn=</c>, which Exchange truncates; that is a guess, and a guess is
    /// not an acceptable basis for deciding whether to send mail.
    /// </remarks>
    private static void AssertOnlyAddressedToSelf(IEnumerable<string?> addresses, OwnerAddresses owner)
    {
        var resolved = addresses.Where(a => !string.IsNullOrWhiteSpace(a)).ToList();

        Assert.NotEmpty(resolved);

        foreach (string? address in resolved)
        {
            bool isOwner = string.Equals(address, owner.Smtp, StringComparison.OrdinalIgnoreCase)
                || (owner.ExchangeDn is not null && string.Equals(address, owner.ExchangeDn, StringComparison.OrdinalIgnoreCase));

            Assert.True(
                isOwner,
                $"Refusing to send: attendee '{address}' is not the mailbox owner ('{owner.Smtp}' / '{owner.ExchangeDn}').");
        }
    }

    /// <summary>
    /// The mailbox owner's own address in both forms Outlook may hand back.
    /// </summary>
    private sealed record OwnerAddresses(string Smtp, string? ExchangeDn);

    private void DeleteCreatedItem(CalendarCommands commands, string? entryId, string? storeId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            return;
        }

        var deleted = commands.DeleteAppointment(entryId: entryId, storeId: storeId, useActiveAppointment: false);

        if (!deleted.Success)
        {
            output.WriteLine($"WARNING: could not delete test meeting {entryId}: {deleted.ErrorMessage}");
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private OwnerAddresses ResolveOwnAddresses()
    {
        OutlookInterop.Application application = RequireOutlook();

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
            string? exchangeDn = entry?.Address;

            exchangeUser = entry?.GetExchangeUser();
            if (exchangeUser is not null)
            {
                address = exchangeUser.PrimarySmtpAddress;
            }

            if (string.IsNullOrWhiteSpace(address) && exchangeDn is string raw && raw.Contains('@', StringComparison.Ordinal))
            {
                address = raw;
            }

            if (string.IsNullOrWhiteSpace(address) || !address.Contains('@', StringComparison.Ordinal))
            {
                throw new SkipException(
                    "Could not resolve the signed-in user's SMTP address; refusing to send anything.");
            }

            return new OwnerAddresses(address, exchangeDn);
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
            Release(exchangeUser);
            Release(entry);
            Release(currentUser);
            Release(session);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private void EnsureOutlookAvailable()
    {
        OutlookInterop.Application application = RequireOutlook();
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
            Release(session);
        }
    }

    private OutlookInterop.Application RequireOutlook()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            output.WriteLine("Skipping self-invite test: no running classic Outlook desktop instance is available.");
            throw new SkipException("No running classic Outlook desktop instance is available.");
        }

        return application!;
    }

    private static void Release(object? comObject)
    {
        if (comObject is not null && System.Runtime.InteropServices.Marshal.IsComObject(comObject))
        {
            _ = System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject);
        }
    }
}
