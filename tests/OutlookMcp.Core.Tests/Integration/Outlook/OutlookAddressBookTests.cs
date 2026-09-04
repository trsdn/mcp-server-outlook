using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.AddressBook;
using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Address book lookup and recipient resolution (#15).
///
/// <para>
/// The failure this surface exists to prevent is discovering at send time that a name did not
/// resolve. Until now the only way to find out whether "Jane Smith" was a real addressee was to
/// build a draft, send it, and wait for a bounce - and Outlook makes that worse than it sounds,
/// because any SMTP-shaped string resolves as a one-off whether or not the mailbox exists.
/// </para>
///
/// <para>
/// The correctness question these tests actually turn on is X500 versus SMTP. An Exchange
/// <c>AddressEntry.Address</c> is a legacyExchangeDN - <c>/o=ExchangeLabs/ou=.../cn=...</c> - and
/// not an email address. Code that reports it as one produces a string that looks plausible,
/// serialises cleanly and is useless to every caller. <see cref="Resolve_ExchangeUser_ReportsSmtpAddressNotTheX500Dn"/>
/// is the test that has to hold.
/// </para>
///
/// <para>
/// Everything here is read-only: no item is created, changed, sent or deleted. Recipients and
/// address entries are, however, prime Object Model Guard territory. Where the guard denies a
/// call, these tests skip with the guard's own message rather than softening an assertion, so a
/// denial is visible in the run rather than reported as a pass.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "AddressBook")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookAddressBookTests(ITestOutputHelper output)
{
    /// <summary>
    /// A profile always has at least one address book, and every one of them must arrive with the
    /// name a caller has to pass back to address it, plus a type it can reason about.
    /// </summary>
    [SkippableFact]
    public void ListAddressLists_ReturnsEveryAddressListWithANameAndAType()
    {
        EnsureOutlookAvailable();

        var result = new AddressBookCommands().ListAddressLists();

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.NotEmpty(result.AddressLists);
        Assert.Equal(result.AddressLists.Count, result.Count);

        foreach (var list in result.AddressLists)
        {
            Assert.False(string.IsNullOrWhiteSpace(list.Name), "An address list arrived without a name.");
            Assert.False(
                string.IsNullOrWhiteSpace(list.AddressListType),
                $"Address list '{list.Name}' arrived without a type.");
            output.WriteLine($"{list.Index}: {list.Name} | type={list.AddressListType} | initial={list.IsInitialAddressList}");
        }
    }

    /// <summary>
    /// The type must be a name, never a raw enum number. <c>olExchangeGlobalAddressList</c> is 0,
    /// so a numeric projection would make the GAL indistinguishable from a falsy default.
    /// </summary>
    [SkippableFact]
    public void ListAddressLists_ReportsTypesAsNamesNotEnumNumbers()
    {
        EnsureOutlookAvailable();

        var result = new AddressBookCommands().ListAddressLists();

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);

        foreach (var list in result.AddressLists)
        {
            Assert.False(
                int.TryParse(list.AddressListType, out _),
                $"Address list '{list.Name}' reported its type as the number '{list.AddressListType}'.");
        }
    }

    /// <summary>
    /// Exactly one address list is the one Outlook opens by default. Reporting none would leave a
    /// caller with no defensible choice of which book to search first.
    /// </summary>
    [SkippableFact]
    public void ListAddressLists_MarksAtMostOneListAsTheInitialOne()
    {
        EnsureOutlookAvailable();

        var result = new AddressBookCommands().ListAddressLists();

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);

        int initial = result.AddressLists.Count(l => l.IsInitialAddressList);
        Assert.True(initial <= 1, $"{initial} address lists claim to be the initial one.");
    }

    /// <summary>
    /// The core correctness test. Resolving the profile's own Exchange address must yield an SMTP
    /// address, not the X500 legacyExchangeDN that <c>AddressEntry.Address</c> hands back for an
    /// Exchange entry.
    /// </summary>
    [SkippableFact]
    public void Resolve_ExchangeUser_ReportsSmtpAddressNotTheX500Dn()
    {
        EnsureOutlookAvailable();

        string address = RequireOwnSmtpAddress();

        var result = new AddressBookCommands().Resolve(address);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);

        var recipient = Assert.Single(result.Recipients);
        Assert.True(recipient.Resolved, $"Outlook could not resolve '{address}', which is this profile's own address.");
        Assert.NotNull(recipient.SmtpAddress);
        Assert.Contains('@', recipient.SmtpAddress);
        Assert.DoesNotContain("/o=", recipient.SmtpAddress, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("/cn=", recipient.SmtpAddress, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(address, recipient.SmtpAddress, ignoreCase: true);

        output.WriteLine(
            $"'{address}' -> smtp={recipient.SmtpAddress} via {recipient.SmtpAddressSource} "
            + $"| entryType={recipient.EntryType} | rawAddress={Truncate(recipient.RawAddress)}");
    }

    /// <summary>
    /// The raw provider address is reported alongside the SMTP one rather than instead of it, so a
    /// caller can see for itself that the two are different things. On an Exchange entry the raw
    /// address is the X500 DN, which is exactly what must never be passed off as an email address.
    /// </summary>
    [SkippableFact]
    public void Resolve_ExchangeUser_ReportsTheRawProviderAddressSeparately()
    {
        EnsureOutlookAvailable();

        string address = RequireOwnSmtpAddress();

        var result = new AddressBookCommands().Resolve(address);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);

        var recipient = Assert.Single(result.Recipients);
        Skip.If(recipient.RawAddress is null, "Outlook declined to report the raw provider address for this entry.");
        Skip.If(
            !string.Equals(recipient.AddressType, "EX", StringComparison.OrdinalIgnoreCase),
            $"This profile resolved its own address as '{recipient.AddressType}', not an Exchange entry.");

        Assert.StartsWith("/o=", recipient.RawAddress, StringComparison.OrdinalIgnoreCase);
        Assert.NotEqual(recipient.RawAddress, recipient.SmtpAddress);
        Assert.Equal("exchange-user", recipient.SmtpAddressSource);
    }

    /// <summary>
    /// A name nobody has is a legitimate answer, not a failure. The operation succeeded; it simply
    /// found nothing. Reporting <c>success: false</c> here would make "I could not reach Outlook"
    /// indistinguishable from "that person does not exist", which is the whole point of the call.
    /// </summary>
    [SkippableFact]
    public void Resolve_UnknownName_SucceedsAndReportsUnresolvedRatherThanFailing()
    {
        EnsureOutlookAvailable();

        string nobody = $"OutlookMcpNoSuchPerson{Guid.NewGuid():N}";

        var result = new AddressBookCommands().Resolve(nobody);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);

        var recipient = Assert.Single(result.Recipients);
        Assert.False(recipient.Resolved, $"Outlook claims to have resolved the invented name '{nobody}'.");
        Assert.Null(recipient.SmtpAddress);
        Assert.False(result.AllResolved);
        Assert.Equal(0, result.ResolvedCount);
        Assert.Contains(nobody, result.UnresolvedNames);
    }

    /// <summary>
    /// The pre-send check has to survive a mixed list. One bad addressee among good ones must be
    /// named individually, not collapse the whole answer into a failure.
    /// </summary>
    [SkippableFact]
    public void Resolve_MixedList_ReportsEachAddresseeIndividually()
    {
        EnsureOutlookAvailable();

        string address = RequireOwnSmtpAddress();
        string nobody = $"OutlookMcpNoSuchPerson{Guid.NewGuid():N}";

        var result = new AddressBookCommands().Resolve($"{address}; {nobody}");

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);

        Assert.Equal(2, result.Recipients.Count);
        Assert.Equal(2, result.RequestedCount);
        Assert.Equal(1, result.ResolvedCount);
        Assert.False(result.AllResolved);

        var good = result.Recipients.Single(r => r.Query.Equals(address, StringComparison.OrdinalIgnoreCase));
        var bad = result.Recipients.Single(r => r.Query == nobody);

        Assert.True(good.Resolved);
        Assert.NotNull(good.SmtpAddress);
        Assert.False(bad.Resolved);
        Assert.Null(bad.SmtpAddress);
        Assert.Equal([nobody], result.UnresolvedNames);
    }

    /// <summary>
    /// Nothing to resolve is a bad request, not an empty success. An empty argument almost always
    /// means the caller built the string wrong, and answering "0 of 0 resolved, all good" would
    /// hand back a green light it never checked.
    /// </summary>
    [SkippableFact]
    public void Resolve_WithNoNames_IsRefused()
    {
        EnsureOutlookAvailable();

        var result = new AddressBookCommands().Resolve("   ");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Empty(result.Recipients);
    }

    /// <summary>
    /// Enumerating an address list must report honestly how much of it was examined. A GAL is
    /// unbounded and has no server-side search in the object model, so a short answer is expected -
    /// what is not acceptable is a short answer that looks complete.
    /// </summary>
    [SkippableFact]
    public void ListEntries_ReportsCountsThatAccountForWhatWasScanned()
    {
        EnsureOutlookAvailable();

        var result = new AddressBookCommands().ListEntries(maxCount: 5);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.AddressListName));

        Assert.Equal(result.Entries.Count, result.ReturnedCount);
        Assert.True(
            result.ScannedCount >= result.ReturnedCount,
            $"{result.ScannedCount} entries scanned but {result.ReturnedCount} returned.");
        Assert.True(result.ReturnedCount <= 5, $"maxCount was 5 but {result.ReturnedCount} entries came back.");

        foreach (var entry in result.Entries)
        {
            Assert.False(string.IsNullOrWhiteSpace(entry.DisplayName), "An address entry arrived without a name.");
            Assert.NotNull(entry.AccessDenied);
            output.WriteLine(
                $"{entry.DisplayName} | {entry.EntryType} | smtp={entry.SmtpAddress ?? "(none)"} "
                + $"| denied={(entry.AccessDenied.Count == 0 ? "-" : string.Join(",", entry.AccessDenied))}");
        }

        output.WriteLine(
            $"'{result.AddressListName}': returned {result.ReturnedCount} of {result.ScannedCount} scanned "
            + $"| truncated={result.Truncated} | scanLimitReached={result.ScanLimitReached}");
    }

    /// <summary>
    /// An address list nobody has must be refused by name rather than silently falling back to the
    /// default book, which would answer a question the caller did not ask.
    /// </summary>
    [SkippableFact]
    public void ListEntries_WithUnknownAddressList_IsRefusedRatherThanFallingBack()
    {
        EnsureOutlookAvailable();

        string nobody = $"OutlookMcpNoSuchAddressList{Guid.NewGuid():N}";

        var result = new AddressBookCommands().ListEntries(addressList: nobody, maxCount: 5);

        Assert.False(result.Success);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Contains(nobody, result.ErrorMessage);
        Assert.Empty(result.Entries);
    }

    /// <summary>
    /// The prefix filter has to actually filter. A filter that is accepted and then ignored is
    /// worse than one that is refused, because the caller believes the answer is narrowed.
    /// </summary>
    [SkippableFact]
    public void ListEntries_WithStartsWith_ReturnsOnlyMatchingEntries()
    {
        EnsureOutlookAvailable();

        var unfiltered = new AddressBookCommands().ListEntries(maxCount: 25);
        SkipIfObjectModelGuardDenied(unfiltered);
        Assert.True(unfiltered.Success, unfiltered.ErrorMessage);
        Skip.If(unfiltered.Entries.Count == 0, "The default address list is empty, so there is nothing to filter.");

        string prefix = unfiltered.Entries[0].DisplayName[..1];

        var filtered = new AddressBookCommands().ListEntries(startsWith: prefix, maxCount: 25);

        SkipIfObjectModelGuardDenied(filtered);
        Assert.True(filtered.Success, filtered.ErrorMessage);
        Assert.NotEmpty(filtered.Entries);

        foreach (var entry in filtered.Entries)
        {
            Assert.StartsWith(prefix, entry.DisplayName, StringComparison.OrdinalIgnoreCase);
        }

        output.WriteLine($"prefix '{prefix}': {filtered.ReturnedCount} of {filtered.ScannedCount} scanned matched.");
    }

    /// <summary>
    /// A comma is part of a name, not a separator. <c>Smith, Jane</c> is the canonical Exchange
    /// Global Address List display-name shape, so splitting on commas would take the single most
    /// common form of the exact input this action exists to resolve and turn it into two addressees
    /// that resolve to nothing.
    ///
    /// <para>
    /// The name here is invented on purpose: what is being pinned is the arity, not the lookup. One
    /// query in must mean one addressee out, whether or not anybody by that name exists.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void Resolve_DisplayNameContainingAComma_IsOneAddresseeNotTwo()
    {
        EnsureOutlookAvailable();

        string commaName = $"Smith, Jane{Guid.NewGuid():N}";

        var result = new AddressBookCommands().Resolve(commaName);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);

        Assert.Equal(1, result.RequestedCount);
        var recipient = Assert.Single(result.Recipients);
        Assert.Equal(commaName, recipient.Query);

        output.WriteLine($"'{commaName}' stayed one addressee; resolved={recipient.Resolved}");
    }

    /// <summary>
    /// The separator contract in full: semicolons split, commas do not, and the two combine
    /// without interfering. Two comma-bearing names either side of a semicolon must arrive as
    /// exactly two addressees with their commas intact.
    /// </summary>
    [SkippableFact]
    public void Resolve_SemicolonSeparatesAndCommaDoesNot()
    {
        EnsureOutlookAvailable();

        string first = $"Smith, Jane{Guid.NewGuid():N}";
        string second = $"Jones, Robert{Guid.NewGuid():N}";

        var result = new AddressBookCommands().Resolve($"{first}; {second}");

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);

        Assert.Equal(2, result.RequestedCount);
        Assert.Equal(2, result.Recipients.Count);
        Assert.Equal([first, second], result.Recipients.Select(r => r.Query));
    }

    /// <summary>
    /// The real-Global-Address-List half of the comma story. A directory of any size contains
    /// display names with commas in them; resolving one must give back a single addressee carrying
    /// the whole name, not two fragments.
    ///
    /// <para>
    /// The name is taken from the live address book rather than invented, so this fails if the
    /// separator contract breaks against real data. It skips, with a stated reason, only if the
    /// scanned slice of this profile's address book happens to hold no comma at all.
    /// </para>
    /// </summary>
    [SkippableFact]
    public void Resolve_RealGalDisplayNameWithAComma_StaysOneAddressee()
    {
        EnsureOutlookAvailable();

        var commands = new AddressBookCommands();

        var listing = commands.ListEntries(maxCount: 100, scanLimit: 2000, includeSmtpAddress: true);
        SkipIfObjectModelGuardDenied(listing);
        Assert.True(listing.Success, listing.ErrorMessage);

        var withComma = listing.Entries.FirstOrDefault(
            e => e.DisplayName.Contains(',', StringComparison.Ordinal));

        Skip.If(
            withComma is null,
            $"No display name among the {listing.ScannedCount} address book entries scanned carries "
            + "a comma, so this profile offers no real comma name to resolve.");

        output.WriteLine($"Real address book comma name: '{withComma!.DisplayName}'");

        var result = commands.Resolve(withComma.DisplayName);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);

        // The load-bearing claim: one name in, one addressee out, comma intact. Whether the
        // directory then resolves it is a separate question - an odd or ambiguous name legitimately
        // may not - but it must never be torn into fragments on the way in.
        Assert.Equal(1, result.RequestedCount);
        var recipient = Assert.Single(result.Recipients);
        Assert.Equal(withComma.DisplayName, recipient.Query);

        if (recipient.Resolved && withComma.SmtpAddress != null && recipient.SmtpAddress != null)
        {
            Assert.Equal(withComma.SmtpAddress, recipient.SmtpAddress, ignoreCase: true);
        }

        output.WriteLine($"resolved={recipient.Resolved}, smtp={recipient.SmtpAddress ?? "(none)"}");
    }

    /// <summary>
    /// A refusal is not an absence. The point of this whole surface is validating an addressee
    /// before a send, and "Outlook has no such person" and "Outlook refused to tell me" call for
    /// opposite actions - correct the address, or treat the answer as unknown and do not call the
    /// send validated. So whenever a resolved addressee comes back with no SMTP address, the reason
    /// must be recoverable: either something is named in <c>accessDenied</c>, or the note says the
    /// directory genuinely holds none.
    /// </summary>
    [SkippableFact]
    public void Resolve_WhenAnAddressIsMissing_SaysWhetherItWasRefusedOrAbsent()
    {
        EnsureOutlookAvailable();

        string address = RequireOwnSmtpAddress();

        var result = new AddressBookCommands().Resolve(address);

        SkipIfObjectModelGuardDenied(result);
        Assert.True(result.Success, result.ErrorMessage);

        var recipient = Assert.Single(result.Recipients);
        Assert.NotNull(recipient.AccessDenied);

        if (recipient is { Resolved: true, SmtpAddress: null })
        {
            Assert.False(
                string.IsNullOrWhiteSpace(recipient.Note),
                "An addressee resolved with no address and no explanation of why.");
        }

        // Nothing on this machine was refused, so the list must be empty rather than merely
        // present - an accessDenied that is never populated would pass the assertion above while
        // telling a caller nothing.
        Assert.Empty(recipient.AccessDenied);
        output.WriteLine(
            $"accessDenied={recipient.AccessDenied.Count}, smtp={recipient.SmtpAddress}, "
            + $"note={recipient.Note ?? "(none)"}");
    }

    /// <summary>
    /// The profile's own delivery address, which is the one addressee guaranteed to exist in this
    /// mailbox's address book. Read through <c>folder.list-stores</c> rather than
    /// <c>NameSpace.CurrentUser</c>, which is itself Object Model Guard protected.
    /// </summary>
    private string RequireOwnSmtpAddress()
    {
        var stores = new FolderCommands().ListStores();
        Skip.If(!stores.Success, stores.ErrorMessage ?? "Stores could not be listed.");

        string? address = stores.Stores
            .Where(s => s.IsDefaultStore)
            .Select(s => s.AccountSmtpAddress)
            .FirstOrDefault(a => !string.IsNullOrWhiteSpace(a))
            ?? stores.Stores
                .Select(s => s.AccountSmtpAddress)
                .FirstOrDefault(a => !string.IsNullOrWhiteSpace(a));

        Skip.If(address is null, "No store in this profile reports an account address to resolve.");
        output.WriteLine($"Resolving against this profile's own address: {address}");
        return address!;
    }

    /// <summary>
    /// Skips rather than failing when Outlook's Object Model Guard refused the call, and says so
    /// in the skip reason. The guard cannot be answered programmatically, so a denial is a fact
    /// about the machine rather than a defect in the code - but it must be visible, not smoothed
    /// over into a pass.
    /// </summary>
    private void SkipIfObjectModelGuardDenied(ResultBase result)
    {
        if (result.Success || result.ErrorMessage is null)
        {
            return;
        }

        if (result.ErrorMessage.Contains("security", StringComparison.OrdinalIgnoreCase)
            || result.ErrorMessage.Contains("Object Model Guard", StringComparison.OrdinalIgnoreCase))
        {
            output.WriteLine($"Object Model Guard denial: {result.ErrorMessage}");
            Skip.If(true, $"Outlook's Object Model Guard denied this call: {result.ErrorMessage}");
        }
    }

    private static string Truncate(string? value)
        => value is null ? "(none)" : value.Length <= 60 ? value : value[..60] + "...";

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
