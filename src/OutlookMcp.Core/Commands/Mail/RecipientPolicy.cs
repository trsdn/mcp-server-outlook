using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// An opt-in allow-list of the recipients Outlook may be asked to send to (#9).
///
/// <para>
/// It covers both outbound paths: <c>mail.send</c>, and <c>calendar.create-appointment</c> with
/// <c>sendInvitation</c>. Those are the two places this surface puts a message addressed to
/// caller-chosen recipients outside the mailbox, and enforcing on only one would be worse than not
/// enforcing at all - the user would believe nothing could leave for an unlisted address while a
/// second route stayed open.
/// </para>
///
/// <para>
/// <b>Off unless configured.</b> With nothing set, <see cref="IsAllowed"/> answers true for
/// everything and send behaves exactly as it did before this existed. That default is deliberate:
/// this server drives the user's own mailbox, and a list nobody asked for would be an obstacle
/// rather than a control.
/// </para>
///
/// <para>
/// <b>Configured through the environment</b>, via <see cref="EnvironmentVariableName"/>. That is
/// the one mechanism the two entry points genuinely share: the MCP server is launched by its
/// client (which sets <c>env</c> in the server definition), and the CLI daemon inherits the shell's
/// environment. Neither has a configuration file that the other reads, and inventing one for this
/// would put the setting in a place only half the product could see.
/// </para>
///
/// <para>
/// The value is a list separated by semicolons or commas. Each entry is either a domain
/// (<c>contoso.example</c> or <c>@contoso.example</c> - both mean the same) or a complete address
/// (<c>alice@contoso.example</c>). Matching is case-insensitive. A domain entry matches that domain
/// exactly and no other: <c>@contoso.example</c> does not admit <c>evilcontoso.example</c>, and
/// does not admit the subdomain <c>mail.contoso.example</c> either - name a subdomain explicitly if
/// it is wanted.
/// </para>
///
/// <para>
/// <b>When enabled, this fails closed.</b> An address the policy cannot parse - empty, malformed,
/// or an unresolved Exchange X500 path rather than SMTP - is refused. The user asked for this
/// control; waving through the addresses it could not read would make it worthless in exactly the
/// cases it exists for. That is the opposite posture from
/// <c>OutlookInteropRunner.IsInDeletedItems</c>, which fails open, and for the opposite reason:
/// there, failing closed would break ordinary work nobody opted into.
/// </para>
/// </summary>
public sealed class RecipientPolicy
{
    /// <summary>
    /// The environment variable read by <see cref="FromEnvironment"/>.
    /// </summary>
    public const string EnvironmentVariableName = "OUTLOOKMCP_ALLOWED_RECIPIENTS";

    private static readonly char[] Separators = [';', ','];

    private readonly List<string> _domains;
    private readonly List<string> _addresses;

    private RecipientPolicy(IReadOnlyList<string> entries, List<string> domains, List<string> addresses)
    {
        Entries = entries;
        _domains = domains;
        _addresses = addresses;
    }

    /// <summary>
    /// The entries exactly as configured, for the refusal message. A refusal that does not say what
    /// the policy permits leaves the caller guessing, and guessing at a security control is how it
    /// ends up switched off.
    /// </summary>
    public IReadOnlyList<string> Entries { get; }

    /// <summary>
    /// True when at least one entry was configured. False means this policy asserts nothing.
    /// </summary>
    public bool IsEnabled => Entries.Count > 0;

    /// <summary>
    /// Reads the policy from <see cref="EnvironmentVariableName"/>.
    /// </summary>
    /// <remarks>
    /// Read per call rather than cached, so changing the variable takes effect on the next send
    /// without restarting the MCP server or the CLI daemon - both of which are long-lived, and
    /// neither of which the user would think to restart after tightening a security setting.
    /// </remarks>
    public static RecipientPolicy FromEnvironment() =>
        Parse(Environment.GetEnvironmentVariable(EnvironmentVariableName));

    public static RecipientPolicy Parse(string? configured)
    {
        var entries = new List<string>();
        var domains = new List<string>();
        var addresses = new List<string>();

        if (!string.IsNullOrWhiteSpace(configured))
        {
            foreach (string raw in configured!.Split(Separators, StringSplitOptions.RemoveEmptyEntries))
            {
                string entry = raw.Trim();
                if (entry.Length == 0)
                {
                    continue;
                }

                entries.Add(entry);

                int at = entry.IndexOf('@');
                if (at > 0)
                {
                    // A local part is present, so this names one mailbox rather than a domain.
                    addresses.Add(entry);
                }
                else
                {
                    // "@contoso.example" and "contoso.example" both mean the domain.
                    domains.Add(at == 0 ? entry[1..] : entry);
                }
            }
        }

        return new RecipientPolicy(entries, domains, addresses);
    }

    /// <summary>
    /// Whether <paramref name="smtpAddress"/> may be sent to. Always true when the policy is
    /// disabled; when enabled, an address that cannot be read as SMTP is refused.
    /// </summary>
    public bool IsAllowed(string? smtpAddress)
    {
        if (!IsEnabled)
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(smtpAddress))
        {
            return false;
        }

        string address = smtpAddress!.Trim();

        int at = address.LastIndexOf('@');
        if (at <= 0 || at == address.Length - 1)
        {
            // Not an SMTP address: a bare name, or an Exchange X500 path that could not be resolved
            // to one. Refused rather than guessed at - see the fail-closed note on the class.
            return false;
        }

        if (_addresses.Any(a => string.Equals(a, address, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        string domain = address[(at + 1)..];
        return _domains.Any(d => string.Equals(d, domain, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// The refusal shown when a send is blocked, naming both what was refused and what is allowed.
    /// </summary>
    public string BuildRefusal(IReadOnlyList<string> rejected)
    {
        string offending = rejected.Count == 0
            ? "one or more recipients whose address could not be read"
            : string.Join(", ", rejected);

        return $"Refused to send: {offending} is outside the configured recipient policy. "
            + $"{EnvironmentVariableName} permits {string.Join(", ", Entries)}. "
            + "This is an opt-in allow-list (#9); an address that cannot be read as SMTP is refused "
            + "rather than assumed safe. Change the recipients, or ask the user to update "
            + $"{EnvironmentVariableName} - do not work around it.";
    }

    /// <summary>
    /// The recipients in <paramref name="recipients"/> this policy does not permit.
    ///
    /// <para>
    /// Shared by <c>mail.send</c> and <c>calendar.create-appointment</c> with
    /// <c>sendInvitation</c>, which are the two paths that put a message addressed to
    /// caller-chosen recipients outside the mailbox. Enforcing on only one of them would be worse
    /// than not enforcing at all: the user would believe the server could not mail outside the list
    /// while a second route stayed open.
    /// </para>
    ///
    /// <para>
    /// <c>Recipient.Address</c> is SMTP for an external recipient and an X500 path
    /// (<c>/o=ExchangeLabs/...</c>) for an internal Exchange one, so an Exchange recipient is
    /// resolved through <c>AddressEntry.GetExchangeUser().PrimarySmtpAddress</c> first. Anything
    /// still not SMTP-shaped is rejected rather than guessed at.
    /// </para>
    ///
    /// <para>
    /// <b>Nothing here catches.</b> <c>Recipients</c> and <c>AddressEntry</c> are Object Model Guard
    /// protected, and a denial must reach <c>OutlookInteropRunner</c>'s classifier so the caller is
    /// told the guard blocked it - not told, falsely, that their recipients failed a policy check.
    /// Rule 1b.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public List<string> CollectDisallowed(Outlook.Recipients recipients)
    {
        var rejected = new List<string>();
        int count = recipients.Count;

        for (int index = 1; index <= count; index++)
        {
            Outlook.Recipient? recipient = null;
            try
            {
                recipient = recipients[index];
                string? address = ResolveSmtpAddress(recipient);

                if (!IsAllowed(address))
                {
                    rejected.Add(string.IsNullOrWhiteSpace(address)
                        ? TryGetName(recipient) ?? $"recipient {index}"
                        : address!);
                }
            }
            finally
            {
                OutlookInteropRunner.ReleaseComObject(ref recipient);
            }
        }

        return rejected;
    }

    /// <summary>
    /// The SMTP address of a recipient, or null when Outlook holds no SMTP form of it.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? ResolveSmtpAddress(Outlook.Recipient recipient)
    {
        string? address = recipient.Address;

        if (!string.IsNullOrWhiteSpace(address) && address!.Contains('@'))
        {
            return address;
        }

        // An internal Exchange recipient: Address is an X500 path, and the SMTP address lives on
        // the ExchangeUser behind the AddressEntry.
        Outlook.AddressEntry? entry = null;
        Outlook.ExchangeUser? exchangeUser = null;
        try
        {
            entry = recipient.AddressEntry;
            exchangeUser = entry?.GetExchangeUser();
            return exchangeUser?.PrimarySmtpAddress ?? address;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref exchangeUser);
            OutlookInteropRunner.ReleaseComObject(ref entry);
        }
    }

    /// <summary>
    /// The recipient's display name, purely to make a refusal readable when no address could be
    /// read. Absence is meaningful and handled: the refusal falls back to the ordinal.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? TryGetName(Outlook.Recipient recipient)
    {
        try
        {
            return recipient.Name;
        }
        catch (COMException)
        {
            return null;
        }
    }
}
