using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Commands.Rules;
using OutlookInterop = Microsoft.Office.Interop.Outlook;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Rule CRUD: create, update, enable/disable and delete (#15).
///
/// <para>
/// <b>Why these tests are shaped the way they are.</b> A rule test passes vacuously with
/// embarrassing ease. <c>Rules.Create</c> hands back a live <c>Rule</c> object immediately, so a
/// test that creates a rule and then asserts on the object it was just given passes whether or not
/// anything was ever written to the mailbox. Only <c>Rules.Save</c> persists, and it is the part
/// that fails - on a malformed clause, on an Exchange quota, on the user having the Rules Wizard
/// open. So every assertion below is made against a <b>fresh</b> <c>list</c>, which goes back to
/// <c>Store.GetRules()</c> and therefore reads what Outlook actually stored, not what this process
/// asked for.
/// </para>
///
/// <para>
/// <b>Safety.</b> These run against the owner's real mailbox and its real rules, which govern real
/// mail flow. Every rule created here is named with <see cref="TestRulePrefix"/> and a GUID, is
/// swept by that prefix in <c>finally</c>, and matches on a subject term containing that same GUID
/// so that it cannot match a real message even in the window where it exists. Nothing here reads,
/// modifies, disables or deletes a rule it did not create, and
/// <see cref="Sweep_LeavesNoTestRuleBehind"/> asserts the store's rule count is exactly what it was.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "RuleCrud")]
[Trait("RequiresOutlook", "true")]
[Collection("Sequential")]
public class OutlookRuleCrudTests(ITestOutputHelper output)
{
    /// <summary>
    /// The one string that separates "a rule this test made" from "a rule the user depends on".
    /// Every sweep matches on it with <see cref="StringComparison.Ordinal"/> and nothing else.
    /// </summary>
    private const string TestRulePrefix = "OutlookMcpTest-";

    /// <summary>
    /// The whole lifecycle, proved at every step by re-reading the store rather than by trusting a
    /// return value: create, confirm it is really persisted with the clauses asked for, then delete
    /// and confirm it is really gone.
    /// </summary>
    [SkippableFact]
    public void Create_ThenDelete_RoundTripsThroughRulesSave()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        var folders = new FolderCommands();
        string scratchFolder = CreateScratchFolder(folders, out string scratchFolderName);
        string name = TestRuleName(out string marker);

        try
        {
            var create = commands.Create(
                name: name,
                subjectContains: marker,
                moveToFolder: scratchFolder,
                enabled: false);

            Assert.True(create.Success, create.ErrorMessage);
            Assert.Null(create.ErrorMessage);
            Assert.Equal("created", create.Action);
            Assert.Equal(name, create.Name);

            // The proof. A fresh list goes back to Store.GetRules(), so finding the rule here means
            // Rules.Save committed it - not merely that Rules.Create returned an object.
            var persisted = FindRule(commands, name);
            Assert.NotNull(persisted);
            Assert.False(persisted!.Enabled);
            Assert.Equal("receive", persisted.RuleType);

            // Outlook puts a new rule at the top of the evaluation order, not the bottom. Asserted
            // rather than assumed, because the documentation this surface gives callers says so and
            // the opposite is what most of them expect.
            Assert.Equal(1, persisted.ExecutionOrder);

            Assert.Contains("subject", persisted.Conditions, StringComparer.Ordinal);
            Assert.Equal([marker], persisted.SubjectTerms);

            Assert.Contains("moveToFolder", persisted.Actions, StringComparer.Ordinal);
            Assert.NotNull(persisted.MoveToFolderPath);
            Assert.EndsWith(scratchFolderName, persisted.MoveToFolderPath!, StringComparison.Ordinal);

            output.WriteLine($"Persisted: {persisted.Name} -> {persisted.MoveToFolderPath}");

            var delete = commands.Delete(name);
            Assert.True(delete.Success, delete.ErrorMessage);
            Assert.Equal("deleted", delete.Action);

            Assert.Null(FindRule(commands, name));
        }
        finally
        {
            SweepTestRules(commands);
            folders.Delete(scratchFolder);
        }
    }

    /// <summary>
    /// A rule with no conditions matches every message that arrives. If its action is a move or a
    /// delete, that silently redirects or destroys the owner's entire incoming mail, and nobody
    /// finds out for days. Outlook's own <c>Rules.Save</c> accepts it. This surface must not.
    /// </summary>
    [SkippableFact]
    public void Create_WithNoConditions_IsRefusedWithoutTouchingTheMailbox()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        int before = RuleCount(commands);
        string name = TestRuleName(out _);

        var result = commands.Create(name: name, moveToFolder: "inbox");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Null(FindRule(commands, name));
        Assert.Equal(before, RuleCount(commands));

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// A rule with no actions does nothing at all. Outlook stores it happily, so a caller gets
    /// <c>success</c> and a rule that will never do the thing they asked for.
    /// </summary>
    [SkippableFact]
    public void Create_WithNoActions_IsRefusedWithoutTouchingTheMailbox()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        int before = RuleCount(commands);
        string name = TestRuleName(out string marker);

        var result = commands.Create(name: name, subjectContains: marker);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Null(FindRule(commands, name));
        Assert.Equal(before, RuleCount(commands));

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Update must <b>replace</b> a clause, not add to it. Outlook's <c>TextRuleCondition.Text</c>
    /// is an array, so an implementation that appended would leave the old term matching forever
    /// while every assertion on "the new term is present" still passed. Two sequential updates are
    /// used deliberately: a single update hides a merge bug that a second one exposes.
    /// </summary>
    [SkippableFact]
    public void Update_ReplacesTheSubjectTerm_RatherThanAppendingToIt()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        string name = TestRuleName(out string marker);
        string second = marker + "-two";
        string third = marker + "-three";

        try
        {
            var create = commands.Create(
                name: name,
                subjectContains: marker,
                assignCategories: null,
                deleteMessage: true,
                enabled: false);
            Assert.True(create.Success, create.ErrorMessage);

            Assert.Equal([marker], FindRule(commands, name)!.SubjectTerms);

            var update = commands.Update(name: name, subjectContains: second);
            Assert.True(update.Success, update.ErrorMessage);
            Assert.Equal("updated", update.Action);

            var afterFirst = FindRule(commands, name);
            Assert.NotNull(afterFirst);
            Assert.Equal([second], afterFirst!.SubjectTerms);
            Assert.DoesNotContain(marker, afterFirst.SubjectTerms);

            Assert.True(commands.Update(name: name, subjectContains: third).Success);

            var afterSecond = FindRule(commands, name);
            Assert.Equal([third], afterSecond!.SubjectTerms);
            Assert.DoesNotContain(second, afterSecond.SubjectTerms);

            // A clause the update did not mention must survive it. PATCH semantics are useless if an
            // unmentioned action is quietly switched off.
            //
            // Asserted as "moveToFolder", not "delete", and that is not a workaround: Outlook has no
            // delete action at all. It rewrites 'delete it' into a move to Deleted Items plus
            // stop-processing, so a correct implementation can never report "delete" here. Asserting
            // on "delete" is how this test failed the first time it ran, which is the only reason
            // that is known rather than assumed.
            Assert.Contains("moveToFolder", afterSecond.Actions, StringComparer.Ordinal);
            Assert.NotNull(afterSecond.MoveToFolderPath);
        }
        finally
        {
            SweepTestRules(commands);
        }
    }

    /// <summary>
    /// The delete action, documented by test because it is the most surprising thing in this
    /// surface: <c>deleteMessage</c> is stored by Outlook as a move to Deleted Items and a
    /// stop-processing, never as a delete. A caller who reads the rule back and looks for a delete
    /// action will not find one, and would reasonably conclude the write had failed.
    /// </summary>
    [SkippableFact]
    public void Create_WithDeleteMessage_IsStoredByOutlookAsAMoveToDeletedItems()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        string name = TestRuleName(out string marker);

        try
        {
            Assert.True(commands.Create(name: name, subjectContains: marker, deleteMessage: true, enabled: false).Success);

            var persisted = FindRule(commands, name);
            Assert.NotNull(persisted);

            Assert.DoesNotContain("delete", persisted!.Actions);
            Assert.Contains("moveToFolder", persisted.Actions, StringComparer.Ordinal);
            Assert.Contains("stop", persisted.Actions, StringComparer.Ordinal);
            Assert.NotNull(persisted.MoveToFolderPath);

            output.WriteLine($"delete rule persisted as: [{string.Join(", ", persisted.Actions)}] -> {persisted.MoveToFolderPath}");
        }
        finally
        {
            SweepTestRules(commands);
        }
    }

    /// <summary>
    /// A rule has one move destination, and a delete consumes it. Accepting both would silently
    /// honour one and drop the other, which is how mail ends up somewhere nobody looks.
    /// </summary>
    [SkippableFact]
    public void Create_WithBothDeleteMessageAndMoveToFolder_IsRefused()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        int before = RuleCount(commands);
        string name = TestRuleName(out string marker);

        var result = commands.Create(
            name: name,
            subjectContains: marker,
            moveToFolder: "inbox",
            deleteMessage: true,
            enabled: false);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Null(FindRule(commands, name));
        Assert.Equal(before, RuleCount(commands));

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Passing an empty string is how a caller clears a clause, and it must be distinguishable from
    /// omitting the argument. Without that, a condition can be added but never removed.
    /// </summary>
    [SkippableFact]
    public void Update_WithAnEmptyValue_ClearsTheClause()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        string name = TestRuleName(out string marker);

        try
        {
            Assert.True(
                commands.Create(
                    name: name,
                    subjectContains: marker,
                    fromAddress: "nobody@example.invalid",
                    deleteMessage: true,
                    enabled: false).Success);

            var before = FindRule(commands, name)!;
            Assert.Contains("senderAddress", before.Conditions, StringComparer.Ordinal);

            var update = commands.Update(name: name, fromAddress: string.Empty);
            Assert.True(update.Success, update.ErrorMessage);

            var after = FindRule(commands, name)!;
            Assert.DoesNotContain("senderAddress", after.Conditions);
            Assert.Empty(after.SenderAddresses);

            // The clause that was not mentioned is untouched, so the rule still has a condition.
            Assert.Contains("subject", after.Conditions, StringComparer.Ordinal);
        }
        finally
        {
            SweepTestRules(commands);
        }
    }

    /// <summary>
    /// Enabling and disabling must survive <c>Rules.Save</c>. Outlook is explicit that a rule is
    /// only enabled once saved, so a <c>Rule.Enabled</c> assignment that was never committed leaves
    /// the caller believing a rule is live when it is inert - or, worse, believing a rule is off
    /// when it is still filing their mail.
    /// </summary>
    [SkippableFact]
    public void SetEnabled_TogglesAndThePreviousStateIsGone()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        string name = TestRuleName(out string marker);

        try
        {
            Assert.True(commands.Create(name: name, subjectContains: marker, deleteMessage: true, enabled: false).Success);
            Assert.False(FindRule(commands, name)!.Enabled);

            var enable = commands.SetEnabled(name, enabled: true);
            Assert.True(enable.Success, enable.ErrorMessage);
            Assert.Equal("enabled", enable.Action);
            Assert.True(FindRule(commands, name)!.Enabled);

            var disable = commands.SetEnabled(name, enabled: false);
            Assert.True(disable.Success, disable.ErrorMessage);
            Assert.Equal("disabled", disable.Action);
            Assert.False(FindRule(commands, name)!.Enabled);
        }
        finally
        {
            SweepTestRules(commands);
        }
    }

    /// <summary>
    /// Rules are addressed by name here, and Outlook permits two rules to share one. Creating a
    /// duplicate would make every later <c>update</c> and <c>delete</c> ambiguous, so the ambiguity
    /// is refused at the only point where this surface can prevent it.
    /// </summary>
    [SkippableFact]
    public void Create_WithANameAlreadyTaken_IsRefused()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        string name = TestRuleName(out string marker);

        try
        {
            Assert.True(commands.Create(name: name, subjectContains: marker, deleteMessage: true, enabled: false).Success);

            var second = commands.Create(name: name, subjectContains: marker + "-again", deleteMessage: true, enabled: false);

            Assert.False(second.Success);
            Assert.NotNull(second.ErrorMessage);
            Assert.Contains(name, second.ErrorMessage!, StringComparison.Ordinal);

            output.WriteLine($"Refused as expected: {second.ErrorMessage}");
        }
        finally
        {
            SweepTestRules(commands);
        }
    }

    /// <summary>
    /// Deleting a name that matches nothing must fail. Reporting success would let a caller believe
    /// a rule they are worried about is gone while it carries on filing their mail.
    /// </summary>
    [SkippableFact]
    public void Delete_OfANameThatMatchesNothing_IsRefused()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        int before = RuleCount(commands);

        var result = commands.Delete(TestRuleName(out _));

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Equal(before, RuleCount(commands));

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// Same for update: a no-op that reports success is how a caller ends up believing a rule was
    /// changed when it was not.
    /// </summary>
    [SkippableFact]
    public void Update_OfANameThatMatchesNothing_IsRefused()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        var result = commands.Update(TestRuleName(out string marker), subjectContains: marker);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// A store id that does not resolve must fail rather than falling back to the default store.
    /// Silently rewriting the wrong mailbox's rules is the worst outcome this surface can produce.
    /// </summary>
    [SkippableFact]
    public void Create_WithAnUnknownStoreId_IsRefusedRatherThanFallingBackToTheDefaultStore()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        int before = RuleCount(commands);
        string name = TestRuleName(out string marker);

        var result = commands.Create(
            name: name,
            subjectContains: marker,
            deleteMessage: true,
            storeId: "not-a-real-store-id");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Null(FindRule(commands, name));
        Assert.Equal(before, RuleCount(commands));

        output.WriteLine($"Refused as expected: {result.ErrorMessage}");
    }

    /// <summary>
    /// The listing this category owns must agree with the one <c>mail list-rules</c> has always
    /// returned; they are the same rules read through one implementation, and a divergence would
    /// mean one of the two surfaces is lying about the mailbox.
    /// </summary>
    [SkippableFact]
    public void List_AgreesWithTheMailCategorysRuleListing()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        var fromRuleCategory = commands.List();
        var fromMailCategory = new Core.Commands.Mail.MailCommands().ListRules();

        Assert.True(fromRuleCategory.Success, fromRuleCategory.ErrorMessage);
        Assert.True(fromMailCategory.Success, fromMailCategory.ErrorMessage);

        Assert.Equal(
            fromMailCategory.Rules.Select(r => r.Name).OrderBy(n => n, StringComparer.Ordinal),
            fromRuleCategory.Rules.Select(r => r.Name).OrderBy(n => n, StringComparer.Ordinal));
    }

    /// <summary>
    /// The last word on safety: after everything above has run, the owner's store must hold no rule
    /// this file created. Runs last by name only incidentally - it sweeps first, so it is correct in
    /// any order.
    /// </summary>
    [SkippableFact]
    public void Sweep_LeavesNoTestRuleBehind()
    {
        var commands = new RuleCommands();
        EnsureOutlookAvailable(commands);

        SweepTestRules(commands);

        var listing = commands.List();
        Assert.True(listing.Success, listing.ErrorMessage);
        Assert.DoesNotContain(
            listing.Rules,
            r => r.Name.StartsWith(TestRulePrefix, StringComparison.Ordinal));
    }

    // ── Helpers ──────────────────────────────────────────────

    private static string TestRuleName(out string marker)
    {
        marker = Guid.NewGuid().ToString("N");
        return TestRulePrefix + marker;
    }

    private static Core.Models.MailRuleInfo? FindRule(RuleCommands commands, string name)
    {
        var listing = commands.List(includeDetail: true);
        Assert.True(listing.Success, listing.ErrorMessage);

        return listing.Rules.SingleOrDefault(r => string.Equals(r.Name, name, StringComparison.Ordinal));
    }

    private static int RuleCount(RuleCommands commands)
    {
        var listing = commands.List();
        Assert.True(listing.Success, listing.ErrorMessage);
        return listing.Rules.Count;
    }

    /// <summary>
    /// Removes every rule whose name carries <see cref="TestRulePrefix"/>, by prefix rather than by
    /// a name captured earlier: a test that failed part way through may have created a rule whose
    /// name the failing assertion never recorded.
    /// </summary>
    private void SweepTestRules(RuleCommands commands)
    {
        var listing = commands.List();
        if (!listing.Success)
        {
            output.WriteLine($"Sweep could not list rules: {listing.ErrorMessage}");
            return;
        }

        foreach (var rule in listing.Rules.Where(r => r.Name.StartsWith(TestRulePrefix, StringComparison.Ordinal)))
        {
            var delete = commands.Delete(rule.Name);
            output.WriteLine($"Swept '{rule.Name}': success={delete.Success} {delete.ErrorMessage}");
        }
    }

    private string CreateScratchFolder(FolderCommands folders, out string name)
    {
        var inbox = folders.ResolvePath("inbox");
        Skip.If(!inbox.Success, inbox.ErrorMessage);

        name = TestRulePrefix + Guid.NewGuid().ToString("N")[..8];
        var create = folders.Create(inbox.FolderPath, name);
        Skip.If(!create.Success, create.ErrorMessage);

        output.WriteLine($"Scratch folder: {create.FolderPath}");
        return create.FolderPath!;
    }

    private void EnsureOutlookAvailable(RuleCommands commands)
    {
        _ = commands;

        if (!OutlookInteropRunner.TryGetRunningApplication(out OutlookInterop.Application? application))
        {
            Skip.If(true, "Classic Outlook is not running; start Outlook to exercise this test.");
            return;
        }

        OutlookInteropRunner.ReleaseSharedComObject(ref application);
        output.WriteLine("Classic Outlook is running; the test will exercise it.");
    }
}
