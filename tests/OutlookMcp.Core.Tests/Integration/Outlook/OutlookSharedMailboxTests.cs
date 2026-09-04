using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Shared and delegate mailbox access (#38, second slice).
///
/// <para>
/// <c>NameSpace.GetSharedDefaultFolder</c> reaches another person's Inbox or Calendar when they have
/// granted access, without that mailbox having to be added to the profile. Nothing in this surface
/// could do that: a user could see their own mailbox and any store already open, and nothing else.
/// </para>
///
/// <para>
/// <b>Testing this honestly is the difficulty.</b> Whether the developer's account has delegate rights
/// over anyone else's mailbox is not something a test can assume, and a test that skips itself when
/// it does not is a test that quietly stops testing. So the exercise here is the user's <i>own</i>
/// address: resolving it and asking for its shared default folder drives exactly the same code path -
/// <c>CreateRecipient</c>, <c>Resolve</c>, <c>GetSharedDefaultFolder</c> - and has a known-correct
/// answer to check against, namely the same folder <c>list-default</c> already returns. Access to a
/// genuinely foreign mailbox is the one part that remains unproven, and the tests say so rather than
/// implying otherwise.
/// </para>
///
/// <para>
/// Everything here is read-only. Nothing is created, moved, sent or deleted.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Store")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookSharedMailboxTests(ITestOutputHelper output)
{
    /// <summary>
    /// The whole path end to end, against the one mailbox access is guaranteed for: the user's own.
    /// The folder returned must be the same one the ordinary default listing gives, otherwise the
    /// shared path is reaching somewhere else and would do so silently for a real delegate too.
    /// </summary>
    [SkippableFact]
    public void OpenShared_ForTheSignedInUser_ReturnsTheirOwnInbox()
    {
        string address = ResolveOwnAddress();

        var commands = new FolderCommands();

        var shared = commands.OpenShared(address, "inbox");

        Assert.True(shared.Success, shared.ErrorMessage);
        Assert.True(shared.Resolved);
        Assert.NotNull(shared.FolderPath);

        var own = commands.ListDefault();
        Assert.True(own.Success, own.ErrorMessage);

        string? ownInbox = own.Folders.Single(f => f.Role == "inbox").FolderPath;

        Assert.Equal(ownInbox, shared.FolderPath);

        output.WriteLine($"Shared inbox for {address} resolved to {shared.FolderPath}");
    }

    /// <summary>
    /// The path handed back has to be usable, or the operation has only told the caller that a
    /// mailbox exists - which they already knew.
    /// </summary>
    [SkippableFact]
    public void OpenShared_ReturnsAPathThatResolvesBack()
    {
        string address = ResolveOwnAddress();

        var commands = new FolderCommands();

        var shared = commands.OpenShared(address, "calendar");
        Assert.True(shared.Success, shared.ErrorMessage);

        var resolved = commands.ResolvePath(shared.FolderPath, includeItemCount: false);

        Assert.True(
            resolved.Success && resolved.Resolved,
            $"'{shared.FolderPath}' was returned by open-shared but resolve-path could not find it: "
            + resolved.ErrorMessage);

        output.WriteLine($"Shared calendar path round-tripped: {shared.FolderPath}");
    }

    /// <summary>
    /// An address nobody answers to must fail rather than return the caller's own mailbox.
    ///
    /// <para>
    /// The trap has two halves. Outlook's <c>Resolve</c> returns false rather than throwing, and an
    /// unresolved recipient passed to <c>GetSharedDefaultFolder</c> is documented to fall back to the
    /// current user's own folder - so a caller asking for a colleague's calendar would silently be
    /// shown their own, with <c>success: true</c>.
    /// </para>
    ///
    /// <para>
    /// The second half was found by running this: <c>Resolve</c> returns <b>true</b> for this
    /// address, because Outlook accepts any syntactically valid SMTP address as a one-off recipient
    /// without consulting the directory. So the resolve guard alone is not enough, and what actually
    /// has to hold is the property asserted here - <b>no folder comes back</b> - regardless of which
    /// layer refuses.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void OpenShared_WithAnUnresolvableAddress_IsRefusedRatherThanReturningOwnFolder()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().OpenShared("no-such-person-9f3a@invalid.example", "inbox");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Null(result.FolderPath);
        Assert.False(result.Resolved);

        // The address must appear in the message. An error that does not name the mailbox it failed
        // on is indistinguishable from a failure about the caller's own.
        Assert.Contains("no-such-person-9f3a@invalid.example", result.ErrorMessage!, StringComparison.Ordinal);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// A role that has no shared equivalent must be refused up front. Outlook accepts only a handful
    /// of folder types here, and passing one it does not accept produces a COM error a caller cannot
    /// act on.
    /// </summary>
    [SkippableFact]
    public void OpenShared_WithAnUnsupportedFolderRole_IsRefusedWithTheSupportedList()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().OpenShared("someone@example.com", "outbox");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("inbox", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// A missing address must be refused rather than defaulting to the current user. "Whose mailbox"
    /// is the entire question this operation answers.
    /// </summary>
    [SkippableFact]
    public void OpenShared_WithNoAddress_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new FolderCommands().OpenShared(null, "inbox");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Reads the signed-in user's own SMTP address. Skips rather than guessing: an address this test
    /// invented would make every assertion above meaningless.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private string ResolveOwnAddress()
    {
        EnsureOutlookAvailable();

        OutlookInterop.Application? application = null;
        OutlookInterop.NameSpace? session = null;
        OutlookInterop.Recipient? currentUser = null;
        OutlookInterop.AddressEntry? entry = null;
        OutlookInterop.ExchangeUser? exchangeUser = null;

        try
        {
            if (!OutlookInteropRunner.TryGetRunningApplication(out application) || application is null)
            {
                throw new SkipException("Classic Outlook is not running.");
            }

            session = application.GetNamespace("MAPI");
            currentUser = session.CurrentUser;
            entry = currentUser?.AddressEntry;

            string? address = null;

            exchangeUser = entry?.GetExchangeUser();
            if (exchangeUser is not null)
            {
                address = exchangeUser.PrimarySmtpAddress;
            }

            if (string.IsNullOrWhiteSpace(address)
                && entry?.Address is string raw
                && raw.Contains('@', StringComparison.Ordinal))
            {
                address = raw;
            }

            if (string.IsNullOrWhiteSpace(address) || !address.Contains('@', StringComparison.Ordinal))
            {
                throw new SkipException("Could not resolve the signed-in user's SMTP address.");
            }

            return address;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref exchangeUser);
            OutlookInteropRunner.ReleaseComObject(ref entry);
            OutlookInteropRunner.ReleaseComObject(ref currentUser);
            OutlookInteropRunner.ReleaseComObject(ref session);
            OutlookInteropRunner.ReleaseSharedComObject(ref application);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private void EnsureOutlookAvailable()
    {
        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
