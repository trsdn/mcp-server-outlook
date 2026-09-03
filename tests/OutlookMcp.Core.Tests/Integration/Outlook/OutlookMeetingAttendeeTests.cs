using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Attendees on calendar items (#32).
///
/// <para>
/// <c>calendar.create-appointment</c> could only ever produce a solo appointment: there was no way to
/// name anybody, and <c>calendar.read</c> never said whether an item was a meeting or who was on it.
/// An agent asked to "set up a call with Anna" would create an entry in the caller's own calendar,
/// report success, and never tell Anna.
/// </para>
///
/// <para>
/// <b>Nothing here sends an invitation.</b> Attendees are attached and resolved, and the item is
/// saved to the caller's own calendar; <c>sendInvitation</c> stays false throughout. The only
/// attendee ever used is the signed-in user, so even a bug that sent something could only reach the
/// mailbox owner. Every created item is deleted in a <c>finally</c>.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CalendarAttendees")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookMeetingAttendeeTests(ITestOutputHelper output)
{
    /// <summary>
    /// The gap itself: naming an attendee must turn the appointment into a meeting and record who is
    /// on it. A saved-but-unsent meeting is the safe half of the feature.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_WithAttendee_BecomesMeetingAndRecordsAttendee()
    {
        string own = ResolveOwnSmtpAddress();
        string subject = $"OutlookMcp attendee test {Guid.NewGuid():N}";
        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(3).Date.AddHours(9);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            requiredAttendees: own);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);
            Assert.True(created.IsMeeting);
            Assert.False(created.InvitationSent);
            Assert.Empty(created.UnresolvedAttendees);

            output.WriteLine($"Created meeting '{subject}' with {created.Attendees.Count} attendee(s), not sent.");
            Assert.Contains(created.Attendees, a => a.Type == "required");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Reading it back must report the same thing. A field that is only correct on the response to
    /// the call that set it has verified nothing about what Outlook actually stored.
    /// </summary>
    [SkippableFact]
    public void Read_ReportsMeetingAndAttendees()
    {
        string own = ResolveOwnSmtpAddress();
        string subject = $"OutlookMcp attendee read {Guid.NewGuid():N}";
        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(4).Date.AddHours(14);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("o"),
            endTime: start.AddMinutes(45).ToString("o"),
            requiredAttendees: own);

        try
        {
            Assert.True(created.Success, created.ErrorMessage);

            var read = commands.Read(entryId: created.EntryId, storeId: created.StoreId, useActiveAppointment: false);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.True(read.HasItem);
            Assert.Equal(subject, read.Subject);
            Assert.True(read.IsMeeting);
            Assert.NotEmpty(read.Attendees);

            foreach (var attendee in read.Attendees)
            {
                output.WriteLine($"  {attendee.Type}: {attendee.Name} <{attendee.Address}> -> {attendee.ResponseStatus}");
            }
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// An attendee Outlook cannot resolve is the dangerous case. Outlook will happily save the
    /// meeting anyway, so a naive implementation reports success for a meeting that can never reach
    /// the person the caller named. It must fail, and it must name who could not be resolved.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_WithUnresolvableAttendee_FailsAndNamesThem()
    {
        EnsureOutlookAvailable();

        string bogus = $"no-such-person-{Guid.NewGuid():N}";
        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(5).Date.AddHours(11);

        var created = commands.CreateAppointment(
            subject: $"OutlookMcp unresolved attendee {Guid.NewGuid():N}",
            start: start.ToString("o"),
            endTime: start.AddMinutes(30).ToString("o"),
            requiredAttendees: bogus);

        try
        {
            Assert.False(created.Success);
            Assert.False(created.InvitationSent);
            Assert.Contains(bogus, created.UnresolvedAttendees);
            Assert.NotNull(created.ErrorMessage);
            Assert.Contains(bogus, created.ErrorMessage!, StringComparison.Ordinal);

            output.WriteLine($"Rejected as expected: {created.ErrorMessage}");
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Without attendees nothing changes. An ordinary appointment must not silently become a meeting.
    /// </summary>
    [SkippableFact]
    public void CreateAppointment_WithoutAttendees_StaysAPlainAppointment()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        DateTimeOffset start = DateTimeOffset.Now.AddDays(6).Date.AddHours(8);

        var created = commands.CreateAppointment(
            subject: $"OutlookMcp solo appointment {Guid.NewGuid():N}",
            start: start.ToString("o"),
            endTime: start.AddMinutes(15).ToString("o"));

        try
        {
            Assert.True(created.Success, created.ErrorMessage);
            Assert.False(created.IsMeeting);
            Assert.False(created.InvitationSent);
            Assert.Empty(created.Attendees);
        }
        finally
        {
            DeleteCreatedItem(commands, created.EntryId, created.StoreId);
        }
    }

    /// <summary>
    /// Removes a calendar item created by a test. Failures are reported rather than swallowed: an
    /// item left in the owner's real calendar is a defect in the test, not an acceptable outcome.
    /// </summary>
    private void DeleteCreatedItem(CalendarCommands commands, string? entryId, string? storeId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            return;
        }

        var deleted = commands.DeleteAppointment(entryId: entryId, storeId: storeId, useActiveAppointment: false);

        if (!deleted.Success)
        {
            output.WriteLine($"WARNING: could not delete test calendar item {entryId}: {deleted.ErrorMessage}");
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
                    "Could not resolve the signed-in user's SMTP address; refusing to name anybody else as an attendee.");
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
            output.WriteLine("Skipping calendar attendee test: no running classic Outlook desktop instance is available.");
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
