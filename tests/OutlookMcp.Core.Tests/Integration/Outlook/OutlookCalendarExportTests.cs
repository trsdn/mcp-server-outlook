using System.Globalization;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Calendar item export - <c>AppointmentItem.SaveAs</c> (#14).
///
/// <para>
/// The mail side of export refuses <c>.ics</c>, because Outlook answers a mail item asked for iCal
/// with "Value does not fall within the expected range" - a message that reads like an argument bug
/// rather than "mail is not a calendar entry". This is the other half: an appointment is the item
/// that <em>can</em> produce iCalendar, and that is the only reason <c>.ics</c> exists in the format
/// table at all.
/// </para>
///
/// <para>
/// <b>Mutation safety.</b> Each test creates its own GUID-named appointment, exports that, and
/// deletes it in <c>finally</c>. <c>sendInvitation</c> is never set and no attendee is ever added,
/// so nothing leaves the machine. Files are written only inside a per-test temp directory that is
/// removed afterwards.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Export")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookCalendarExportTests(ITestOutputHelper output)
{
    /// <summary>
    /// The point of the calendar side: an appointment exported as <c>.ics</c> must be real
    /// iCalendar that another calendar application would accept, not merely a file that exists.
    /// </summary>
    [SkippableFact]
    public void Export_AppointmentToIcs_WritesRealICalendar()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        using var scratch = new ScratchDirectory();
        string subject = ScratchSubject();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = CreateScratchAppointment(commands, subject);
            entryId = created.EntryId;
            storeId = created.StoreId;

            string path = Path.Combine(scratch.Path, "appointment.ics");
            var exported = commands.Export(
                path,
                entryId: entryId,
                storeId: storeId,
                useActiveAppointment: false);

            Assert.True(exported.Success, exported.ErrorMessage);
            Assert.Equal("ics", exported.Format);
            Assert.True(exported.BytesWritten > 0);
            Assert.Equal(new FileInfo(path).Length, exported.BytesWritten);

            string ical = File.ReadAllText(path);
            Assert.StartsWith("BEGIN:VCALENDAR", ical, StringComparison.Ordinal);
            Assert.Contains("BEGIN:VEVENT", ical, StringComparison.Ordinal);
            Assert.Contains("END:VCALENDAR", ical, StringComparison.Ordinal);
            Assert.Contains(subject, ical, StringComparison.Ordinal);

            output.WriteLine($"wrote {exported.BytesWritten} bytes of iCalendar");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// An appointment is not only iCalendar. The same path rules and the same Unicode-safe msg
    /// mapping apply here, so exporting an appointment as <c>.msg</c> must produce the OLE compound
    /// file that a <c>.msg</c> is.
    /// </summary>
    [SkippableFact]
    public void Export_AppointmentToMsg_WritesACompoundFile()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        using var scratch = new ScratchDirectory();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = CreateScratchAppointment(commands, ScratchSubject());
            entryId = created.EntryId;
            storeId = created.StoreId;

            string path = Path.Combine(scratch.Path, "appointment.msg");
            var exported = commands.Export(
                path,
                entryId: entryId,
                storeId: storeId,
                useActiveAppointment: false);

            Assert.True(exported.Success, exported.ErrorMessage);
            Assert.Equal("msg", exported.Format);

            byte[] header = new byte[4];
            using (var stream = File.OpenRead(path))
            {
                Assert.Equal(4, stream.Read(header, 0, 4));
            }

            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, header);

            output.WriteLine($"wrote {exported.BytesWritten} bytes of msg");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// The destination rules are shared with mail export, so they must hold here too. A relative
    /// path resolves against Outlook's working directory rather than the caller's, and an export
    /// that lands somewhere invisible is worse than one that fails.
    /// </summary>
    [SkippableFact]
    public void Export_AppointmentRefusesARelativePath()
    {
        EnsureOutlookAvailable();

        var commands = new CalendarCommands();
        string? entryId = null;
        string? storeId = null;

        try
        {
            var created = CreateScratchAppointment(commands, ScratchSubject());
            entryId = created.EntryId;
            storeId = created.StoreId;

            var exported = commands.Export(
                "relative-appointment-probe.ics",
                entryId: entryId,
                storeId: storeId,
                useActiveAppointment: false);

            Assert.False(exported.Success);
            Assert.Contains("absolute", exported.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

            output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
        }
        finally
        {
            DeleteIfCreated(commands, entryId, storeId);
        }
    }

    /// <summary>
    /// An id that resolves to nothing must not leave a file behind. An export that writes an empty
    /// or wrong item and reports success is worse than one that fails.
    /// </summary>
    [SkippableFact]
    public void Export_AppointmentRefusesAnIdThatDoesNotResolveAndWritesNothing()
    {
        EnsureOutlookAvailable();

        using var scratch = new ScratchDirectory();
        string path = Path.Combine(scratch.Path, "never.ics");

        var exported = new CalendarCommands().Export(
            path,
            entryId: "0000000000000000000000000000000000000000000000",
            useActiveAppointment: false);

        Assert.False(exported.Success);
        Assert.False(File.Exists(path), "A failed export still wrote a file.");

        output.WriteLine($"Refused as expected: {exported.ErrorMessage}");
    }

    private static (string? EntryId, string? StoreId) CreateScratchAppointment(
        CalendarCommands commands,
        string subject)
    {
        // Well into the future, so a scratch appointment cannot collide with anything the owner is
        // actually doing, and no reminder can fire during the run.
        DateTime start = DateTime.Today.AddYears(1).AddHours(9);

        var created = commands.CreateAppointment(
            subject: subject,
            start: start.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
            endTime: start.AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
            body: "export probe");

        Assert.True(created.Success, created.ErrorMessage);
        return (created.EntryId, created.StoreId);
    }

    private void DeleteIfCreated(CalendarCommands commands, string? entryId, string? storeId)
    {
        if (string.IsNullOrEmpty(entryId))
        {
            return;
        }

        var deleted = commands.DeleteAppointment(entryId, storeId);
        output.WriteLine($"Delete appointment: success={deleted.Success} {deleted.ErrorMessage}");
    }

    private static string ScratchSubject() => $"mcp-export-test-{Guid.NewGuid():N}";

    private sealed class ScratchDirectory : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"mcp-export-cal-{Guid.NewGuid():N}");

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
