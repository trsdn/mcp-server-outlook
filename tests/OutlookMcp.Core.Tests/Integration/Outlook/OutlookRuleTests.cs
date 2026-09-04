using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Inbox rule discovery (#15).
///
/// <para>
/// Rules silently move, delete and forward mail before anything in this surface ever sees it. Until
/// now they were invisible here, which makes a whole class of question unanswerable: "why is nothing
/// arriving in my inbox from this person?" has an obvious answer that the tool could not see, so it
/// would confidently report an empty folder instead. That is the failure this project keeps finding -
/// a truthful answer to a question the caller did not ask.
/// </para>
///
/// <para>
/// Read-only throughout. Rules are a mailbox-wide setting that governs real mail flow; nothing here
/// creates, enables, disables or deletes one.
/// </para>
///
/// <para>
/// Written not to assume the owner's profile has any particular rules, or any at all. A profile with
/// no rules is legitimate, and a test that only passes on a rule-heavy profile is a test that
/// silently stops testing anything.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "MailRule")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookRuleTests(ITestOutputHelper output)
{
    /// <summary>
    /// The baseline: rules are enumerated, and each arrives with the name the user sees in Outlook
    /// plus the execution order that decides which one wins. Order matters because an earlier rule
    /// can stop processing entirely, so a listing without it cannot explain what actually happened.
    /// </summary>
    [SkippableFact]
    public void ListRules_ReturnsEveryRuleWithNameAndExecutionOrder()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListRules();

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Rules.Count == 0, "This profile has no rules defined.");

        foreach (var rule in result.Rules)
        {
            Assert.False(string.IsNullOrWhiteSpace(rule.Name), "A rule arrived without a name.");
            Assert.True(rule.ExecutionOrder > 0, $"Rule '{rule.Name}' reported execution order {rule.ExecutionOrder}.");
        }

        Assert.Equal(
            result.Rules.Select(r => r.ExecutionOrder).Distinct().Count(),
            result.Rules.Count);

        output.WriteLine($"{result.Rules.Count} rule(s); {result.Rules.Count(r => r.Enabled)} enabled.");
    }

    /// <summary>
    /// Whether a rule runs on arriving or on outgoing mail changes what its existence explains. A
    /// listing that collapsed the two would answer "why did this get moved?" with a send rule.
    /// </summary>
    [SkippableFact]
    public void ListRules_ReportsRuleTypeByNameRatherThanEnumOrdinal()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListRules();

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Rules.Count == 0, "This profile has no rules defined.");

        foreach (var rule in result.Rules)
        {
            Assert.True(
                rule.RuleType is "receive" or "send",
                $"Rule '{rule.Name}' reported ruleType '{rule.RuleType}'.");
        }
    }

    /// <summary>
    /// The trap this test exists for: Outlook's <c>Conditions</c> and <c>Actions</c> collections have
    /// a fixed length - every rule reports the same ~31 conditions and ~28 actions, because those are
    /// the available <i>slots</i>, not the ones in use. Reporting the raw count would tell a caller
    /// that a one-line rule has 31 conditions, which is not merely useless but actively false.
    /// </summary>
    [SkippableFact]
    public void ListRules_ReportsOnlyTheClausesActuallyInUse()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListRules(includeDetail: true);

        Assert.True(result.Success, result.ErrorMessage);
        Skip.If(result.Rules.Count == 0, "This profile has no rules defined.");

        var withDetail = result.Rules.Where(r => r.Conditions.Count > 0 || r.Actions.Count > 0).ToList();
        Skip.If(withDetail.Count == 0, "No rule in this profile has an enabled condition or action.");

        foreach (var rule in withDetail)
        {
            // A rule that genuinely used every slot Outlook offers is not a thing users build; a
            // count at the collection length means the enabled check was skipped.
            Assert.True(
                rule.Conditions.Count < 20,
                $"Rule '{rule.Name}' reported {rule.Conditions.Count} conditions, which is the size of "
                + "Outlook's fixed slot collection rather than the clauses in use.");

            Assert.True(
                rule.Actions.Count < 20,
                $"Rule '{rule.Name}' reported {rule.Actions.Count} actions, which is the size of "
                + "Outlook's fixed slot collection rather than the clauses in use.");

            foreach (var condition in rule.Conditions)
            {
                Assert.False(string.IsNullOrWhiteSpace(condition), $"Rule '{rule.Name}' had a nameless condition.");
                Assert.False(int.TryParse(condition, out _), $"Rule '{rule.Name}' reported condition '{condition}' as a raw ordinal.");
            }

            foreach (var action in rule.Actions)
            {
                Assert.False(string.IsNullOrWhiteSpace(action), $"Rule '{rule.Name}' had a nameless action.");
                Assert.False(int.TryParse(action, out _), $"Rule '{rule.Name}' reported action '{action}' as a raw ordinal.");
            }
        }

        var sample = withDetail[0];
        output.WriteLine($"{sample.Name}: [{string.Join(", ", sample.Conditions)}] -> [{string.Join(", ", sample.Actions)}]");
    }

    /// <summary>
    /// "Where does my mail go?" is the single most useful thing a rule listing can answer, so a
    /// move-to-folder rule must name its destination. Without it the caller knows mail is being moved
    /// and still cannot go and find it.
    /// </summary>
    [SkippableFact]
    public void ListRules_NamesTheDestinationOfAMoveRule()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListRules(includeDetail: true);

        Assert.True(result.Success, result.ErrorMessage);

        var movers = result.Rules
            .Where(r => r.Actions.Contains("moveToFolder", StringComparer.Ordinal))
            .ToList();

        Skip.If(movers.Count == 0, "This profile has no move-to-folder rules.");

        foreach (var rule in movers)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(rule.MoveToFolderPath),
                $"Rule '{rule.Name}' moves mail but did not say where to.");
        }

        output.WriteLine($"{movers.Count} move rule(s), e.g. {movers[0].Name} -> {movers[0].MoveToFolderPath}");
    }

    /// <summary>
    /// Rule recipients are stored unresolved: <c>Recipient.Address</c> is blank and the address sits
    /// in <c>Name</c>. Reading only <c>Address</c> would report every from-rule as matching nobody -
    /// a listing that looks complete and says the opposite of the truth.
    /// </summary>
    [SkippableFact]
    public void ListRules_ReportsTheAddressesAFromRuleMatches()
    {
        EnsureOutlookAvailable();

        var result = new MailCommands().ListRules(includeDetail: true);

        Assert.True(result.Success, result.ErrorMessage);

        var fromRules = result.Rules
            .Where(r => r.Conditions.Contains("from", StringComparer.Ordinal))
            .ToList();

        Skip.If(fromRules.Count == 0, "This profile has no from-address rules.");

        foreach (var rule in fromRules)
        {
            Assert.True(
                rule.FromAddresses.Count > 0,
                $"Rule '{rule.Name}' matches on sender but reported no addresses at all.");

            foreach (string address in rule.FromAddresses)
            {
                Assert.False(string.IsNullOrWhiteSpace(address), $"Rule '{rule.Name}' reported a blank sender.");
            }
        }

        output.WriteLine($"{fromRules.Count} from-rule(s), e.g. {fromRules[0].Name} <- {string.Join(", ", fromRules[0].FromAddresses)}");
    }

    /// <summary>
    /// Detail costs roughly forty times as much COM work as the summary, because it walks every
    /// condition and action slot of every rule. The cheap listing must therefore genuinely skip that
    /// work rather than gathering it and discarding it, or the parameter is a lie.
    /// </summary>
    [SkippableFact]
    public void ListRules_WithoutDetail_DoesNotGatherClauses()
    {
        EnsureOutlookAvailable();

        var commands = new MailCommands();

        var summary = commands.ListRules();
        Assert.True(summary.Success, summary.ErrorMessage);
        Skip.If(summary.Rules.Count == 0, "This profile has no rules defined.");

        Assert.All(summary.Rules, rule =>
        {
            Assert.Empty(rule.Conditions);
            Assert.Empty(rule.Actions);
            Assert.Null(rule.MoveToFolderPath);
            Assert.Empty(rule.FromAddresses);
        });

        var detailed = commands.ListRules(includeDetail: true);
        Assert.True(detailed.Success, detailed.ErrorMessage);

        // Same rules either way - detail adds clauses, it must not change which rules exist.
        Assert.Equal(summary.Rules.Count, detailed.Rules.Count);
        Assert.Equal(
            summary.Rules.Select(r => r.Name).OrderBy(n => n, StringComparer.Ordinal),
            detailed.Rules.Select(r => r.Name).OrderBy(n => n, StringComparer.Ordinal));

        Assert.Contains(detailed.Rules, r => r.Conditions.Count > 0 || r.Actions.Count > 0);
    }

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
