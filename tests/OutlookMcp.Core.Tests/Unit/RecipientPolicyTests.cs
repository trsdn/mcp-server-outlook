using OutlookMcp.Core.Commands.Mail;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Parsing and matching for the opt-in recipient allow-list (#9).
///
/// <para>
/// Pure string logic with no COM involvement: the policy is built from a configuration string and
/// asked about addresses that are already resolved. Extracting those addresses from a real
/// <c>MailItem</c> is a separate, COM-dependent problem covered by
/// <c>OutlookRecipientPolicyTests</c>. Qualifies under the ADR-001 exception.
/// </para>
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "RecipientPolicy")]
[Trait("Speed", "Fast")]
[Trait("RequiresOutlook", "false")]
public class RecipientPolicyTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Parse_WithNothingConfigured_IsDisabled(string? configured)
    {
        var policy = RecipientPolicy.Parse(configured);

        Assert.False(policy.IsEnabled);

        // Disabled must mean "unchanged behaviour", not "allow-list of nothing". Getting this
        // backwards would break every existing send the moment the feature shipped.
        Assert.True(policy.IsAllowed("anyone@anywhere.example"));
    }

    [Fact]
    public void Parse_WithEntries_IsEnabled()
    {
        var policy = RecipientPolicy.Parse("@contoso.example");

        Assert.True(policy.IsEnabled);
    }

    [Theory]
    // A bare domain and an @-prefixed domain mean the same thing.
    [InlineData("@contoso.example", "alice@contoso.example", true)]
    [InlineData("contoso.example", "alice@contoso.example", true)]
    [InlineData("@contoso.example", "alice@fabrikam.example", false)]
    // Case is not significant anywhere in an address.
    [InlineData("@Contoso.Example", "ALICE@contoso.example", true)]
    // A domain entry must not match a subdomain or a lookalike suffix. "evilcontoso.example"
    // ends with "contoso.example" and must be refused.
    [InlineData("@contoso.example", "alice@evilcontoso.example", false)]
    [InlineData("@contoso.example", "alice@mail.contoso.example", false)]
    // An entry with a local part is an exact address, not a domain.
    [InlineData("alice@contoso.example", "alice@contoso.example", true)]
    [InlineData("alice@contoso.example", "bob@contoso.example", false)]
    // Several entries, any of which may match.
    [InlineData("@contoso.example; bob@fabrikam.example", "bob@fabrikam.example", true)]
    [InlineData("@contoso.example, bob@fabrikam.example", "carol@fabrikam.example", false)]
    // Whitespace around entries is the normal shape of a hand-edited setting.
    [InlineData("  @contoso.example ;  bob@fabrikam.example  ", "alice@contoso.example", true)]
    public void IsAllowed_MatchesDomainsAndExactAddresses(string configured, string address, bool expected)
    {
        var policy = RecipientPolicy.Parse(configured);

        Assert.Equal(expected, policy.IsAllowed(address));
    }

    /// <summary>
    /// An address the policy cannot make sense of must be refused, not waved through. This is an
    /// opt-in control the user asked for: failing open here would make it worthless precisely in
    /// the cases it exists for.
    /// </summary>
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not-an-address")]
    [InlineData("/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=abc")]
    public void IsAllowed_WithAnUnusableAddress_RefusesWhenEnabled(string? address)
    {
        var policy = RecipientPolicy.Parse("@contoso.example");

        Assert.False(policy.IsAllowed(address));
    }

    /// <summary>
    /// The same unusable addresses are fine when no policy is configured, because then nothing is
    /// being asserted about them at all.
    /// </summary>
    [Fact]
    public void IsAllowed_WithAnUnusableAddress_IsFineWhenDisabled()
    {
        var policy = RecipientPolicy.Parse(null);

        Assert.True(policy.IsAllowed(null));
        Assert.True(policy.IsAllowed("not-an-address"));
    }

    [Fact]
    public void Parse_IgnoresEmptyEntries()
    {
        var policy = RecipientPolicy.Parse(";;  ;;");

        Assert.False(policy.IsEnabled);
    }

    /// <summary>
    /// The configured entries come back for the error message. A refusal that does not say what the
    /// policy actually permits leaves the caller guessing, and guessing at a security control is
    /// how it gets disabled.
    /// </summary>
    [Fact]
    public void Entries_AreReportedForTheRefusalMessage()
    {
        var policy = RecipientPolicy.Parse("@contoso.example; bob@fabrikam.example");

        Assert.Equal(["@contoso.example", "bob@fabrikam.example"], policy.Entries);
    }
}
