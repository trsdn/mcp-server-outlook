using OutlookMcp.Core.Commands.Mail;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// The tied-received-time boundary a paging cursor is built from (#135).
///
/// <para>
/// This is a permitted pure test under ADR-001: <see cref="MailPageBoundary"/> and
/// <see cref="MailPageCursor"/> touch no COM object at all, and the defect under test is entirely in
/// how ids are accumulated across calls. It proves only that narrow claim - it is not coverage for
/// any Outlook operation, and the listing paths that use it are covered by integration tests.
/// </para>
///
/// <para>
/// The walk below is the bug in miniature, and it is the reason this is tested at the level of the
/// boundary rather than through a folder: three messages must share a received time <b>to the
/// tick</b>, which is easy to state here and very hard to arrange in a real mailbox - a draft this
/// test created would get its own creation instant, so a test built on drafts would compare distinct
/// timestamps, never enter the tied band, and pass without exercising a line of the logic. That is
/// precisely how the defect survived.
/// </para>
/// </summary>
[Trait("Category", "Unit")]
[Trait("Feature", "MailPaging")]
public class MailPageBoundaryTests
{
    private const string Fingerprint = "abc123def4567890";

    /// <summary>An instant shared by every message in the band, to the tick.</summary>
    private static readonly DateTimeOffset Tie = new(2026, 3, 7, 9, 0, 0, TimeSpan.Zero);

    private static readonly DateTimeOffset Older = Tie.AddMinutes(-5);

    /// <summary>
    /// The whole contract: paging one item at a time through three messages that share a received
    /// time must serve each exactly once and then stop.
    ///
    /// <para>
    /// Before the fix this oscillates - A, B, A, B, ... - because each cursor carried only its own
    /// page's ids and forgot the previous page's, so the resume filter re-admitted A. C is never
    /// reached and the walk never ends, while every response reports success.
    /// </para>
    /// </summary>
    [Fact]
    public void Walk_OverMessagesSharingAReceivedTime_ServesEachOnceAndTerminates()
    {
        string[] band = ["A", "B", "C"];

        var served = new List<string>();
        MailPageCursor? cursor = null;

        // Generously more pages than the three the walk should need. A walk that fails to advance
        // exhausts this and fails the assertion below rather than hanging.
        for (int page = 0; page < 12; page++)
        {
            var boundary = new MailPageBoundary(cursor);
            string? emitted = null;

            // One page of a scan: the band is re-read from the top, already-served ids are excluded
            // by identity, and the page stops at its first match (maxCount = 1).
            foreach (string id in band)
            {
                if (cursor != null && !cursor.Includes(Tie, id))
                {
                    continue;
                }

                boundary.Observe(Tie, id);
                emitted = id;
                break;
            }

            if (emitted == null)
            {
                cursor = null;
                break;
            }

            served.Add(emitted);

            string token = MailPageCursor.Encode(Fingerprint, boundary.Instant!.Value, boundary.Ids);
            Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out cursor, out string? error), error);
        }

        Assert.Equal(band, served);
        Assert.Null(cursor);
    }

    /// <summary>
    /// The specific mechanism, stated on its own so a failure points at the cause rather than at the
    /// walk: while the frontier is still the instant the cursor stopped at, the ids that cursor
    /// carried must survive into the next one.
    /// </summary>
    [Fact]
    public void Observe_AtTheCursorsOwnInstant_KeepsTheIdsAlreadyServed()
    {
        MailPageCursor cursor = Decode(Tie, ["A"]);

        var boundary = new MailPageBoundary(cursor);
        boundary.Observe(Tie, "B");

        Assert.Equal(Tie, boundary.Instant);
        Assert.Equal(["A", "B"], boundary.Ids);
    }

    /// <summary>
    /// And the other half: on genuinely advancing to an older instant the previous band is finished
    /// with, so carrying it forward would grow every cursor without bound and exclude ids that were
    /// never at this instant to begin with.
    /// </summary>
    [Fact]
    public void Observe_AtAnOlderInstant_ForgetsThePreviousBand()
    {
        MailPageCursor cursor = Decode(Tie, ["A"]);

        var boundary = new MailPageBoundary(cursor);
        boundary.Observe(Tie, "B");
        boundary.Observe(Older, "C");

        Assert.Equal(Older, boundary.Instant);
        Assert.Equal(["C"], boundary.Ids);
    }

    /// <summary>A first page has nothing to inherit, and must not invent anything.</summary>
    [Fact]
    public void Observe_WithoutACursor_StartsFromTheItemsItActuallySaw()
    {
        var boundary = new MailPageBoundary(null);
        boundary.Observe(Tie, "A");

        Assert.Equal(Tie, boundary.Instant);
        Assert.Equal(["A"], boundary.Ids);
    }

    /// <summary>
    /// An id seen twice must not be recorded twice. It would not corrupt the walk - the resume filter
    /// treats the list as a set - but a long tied band would grow the cursor on every page.
    /// </summary>
    [Fact]
    public void Observe_WithARepeatedId_RecordsItOnce()
    {
        var boundary = new MailPageBoundary(Decode(Tie, ["A"]));
        boundary.Observe(Tie, "A");
        boundary.Observe(Tie, "B");

        Assert.Equal(["A", "B"], boundary.Ids);
    }

    private static MailPageCursor Decode(DateTimeOffset instant, string[] ids)
    {
        string token = MailPageCursor.Encode(Fingerprint, instant, ids);
        Assert.True(MailPageCursor.TryDecode(token, Fingerprint, out MailPageCursor? cursor, out string? error), error);
        return cursor!;
    }
}
