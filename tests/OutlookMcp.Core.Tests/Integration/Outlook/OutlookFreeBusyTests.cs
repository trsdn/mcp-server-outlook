using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Free/busy lookup (#32).
///
/// <para>
/// "Find a slot that works for everyone" was impossible: nothing exposed
/// <c>Recipient.FreeBusy</c>, so an agent asked to schedule around somebody could only guess. These
/// tests query the signed-in user's own availability - reading it requires no permission the caller
/// does not already have, and asks nobody else's server anything.
/// </para>
///
/// <para>
/// They are read-only. Nothing is created, sent or deleted.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "CalendarFreeBusy")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookFreeBusyTests(ITestOutputHelper output)
{
    /// <summary>
    /// The gap itself: asking when somebody is free must return something a caller can act on.
    /// </summary>
    [SkippableFact]
    public void GetFreeBusy_ForSelf_ReturnsAvailability()
    {
        string own = ResolveOwnSmtpAddress();
        DateTimeOffset start = DateTimeOffset.Now.Date;

        var result = new CalendarCommands().GetFreeBusy(
            attendees: own,
            start: start.ToString("o"),
            days: 3);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.People);

        var person = result.People[0];
        Assert.True(person.Resolved);
        Assert.NotNull(person.Availability);
        Assert.NotEmpty(person.Availability!);

        output.WriteLine($"{person.Name}: {person.Availability!.Length} slot(s) of {result.IntervalMinutes} min, {person.BusyPeriods.Count} busy period(s).");

        foreach (var period in person.BusyPeriods.Take(5))
        {
            output.WriteLine($"  {period.Start:g} - {period.End:g}: {period.Status}");
        }
    }

    /// <summary>
    /// The slot string is only useful if its length matches the window it claims to describe.
    /// A caller reading "free" off the end of a short string would schedule into unknown time.
    /// </summary>
    [SkippableFact]
    public void GetFreeBusy_AvailabilityCoversTheRequestedWindow()
    {
        string own = ResolveOwnSmtpAddress();
        DateTimeOffset start = DateTimeOffset.Now.Date;
        const int days = 2;

        var result = new CalendarCommands().GetFreeBusy(
            attendees: own,
            start: start.ToString("o"),
            days: days,
            intervalMinutes: 60);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(60, result.IntervalMinutes);

        int expected = days * 24;
        Assert.Equal(expected, result.People[0].Availability!.Length);
        Assert.Equal(start.AddDays(days), result.End);
    }

    /// <summary>
    /// Every decoded busy period must fall inside the requested window and end after it starts.
    /// A decoder that is off by one slot silently proposes meetings on top of existing ones.
    /// </summary>
    [SkippableFact]
    public void GetFreeBusy_BusyPeriodsStayInsideTheWindow()
    {
        string own = ResolveOwnSmtpAddress();
        DateTimeOffset start = DateTimeOffset.Now.Date;

        var result = new CalendarCommands().GetFreeBusy(
            attendees: own,
            start: start.ToString("o"),
            days: 5,
            intervalMinutes: 30);

        Assert.True(result.Success, result.ErrorMessage);

        foreach (var period in result.People[0].BusyPeriods)
        {
            Assert.True(period.End > period.Start, "A busy period must end after it starts.");
            Assert.True(period.Start >= result.Start, "A busy period must not start before the window.");
            Assert.True(period.End <= result.End, "A busy period must not end after the window.");
            Assert.NotEqual("free", period.Status);
        }
    }

    /// <summary>
    /// An address Outlook cannot resolve must be reported as such. Reporting it as free would put a
    /// meeting on somebody's calendar on the strength of a lookup that never happened.
    /// </summary>
    [SkippableFact]
    public void GetFreeBusy_UnresolvableAttendee_IsReportedNotAssumedFree()
    {
        EnsureOutlookAvailable();

        string bogus = $"no-such-person-{Guid.NewGuid():N}";

        var result = new CalendarCommands().GetFreeBusy(
            attendees: bogus,
            start: DateTimeOffset.Now.Date.ToString("o"),
            days: 1);

        Assert.False(result.Success);
        Assert.Contains(bogus, result.UnresolvedAttendees);
        Assert.Contains(bogus, result.ErrorMessage!, StringComparison.Ordinal);
        Assert.DoesNotContain(result.People, p => p.Resolved && p.Name == bogus);

        output.WriteLine($"Rejected as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Asking about nobody is a caller mistake, not an empty answer.
    /// </summary>
    [SkippableFact]
    public void GetFreeBusy_WithoutAttendees_Fails()
    {
        EnsureOutlookAvailable();

        var result = new CalendarCommands().GetFreeBusy(
            attendees: "   ",
            start: DateTimeOffset.Now.Date.ToString("o"),
            days: 1);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

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
                    "Could not resolve the signed-in user's SMTP address; refusing to query anybody else's availability.");
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
            output.WriteLine("Skipping free/busy test: no running classic Outlook desktop instance is available.");
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
