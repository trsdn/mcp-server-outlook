using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Rules;

/// <summary>
/// Inbox rule management against the user's running Outlook (#15).
///
/// <para>
/// <b>The three things about Outlook's rule model that shape every method here.</b>
/// </para>
///
/// <para>
/// First, nothing persists until <c>Rules.Save</c>. <c>Rules.Create</c> hands back a live
/// <c>Rule</c> object immediately and populating it looks like it worked, but the mailbox is
/// untouched until the save commits. That makes the save the only step whose failure matters, and
/// it fails for reasons that have nothing to do with the caller's rule: the Exchange rules quota is
/// full, the user has the Rules and Alerts wizard open on the same store, or some other rule in the
/// collection is malformed. So every mutation here does all of its validation <em>before</em>
/// touching the collection, and a failed save is reported as "nothing was written" rather than as an
/// ambiguous error.
/// </para>
///
/// <para>
/// Second, the save is collection-wide. There is no per-rule write: changing one rule rewrites all
/// of a store's rules. That is why every mutation reports <c>ruleCount</c>.
/// </para>
///
/// <para>
/// Third, rules belong to a store, not to a session. <c>Store.GetRules</c> is the only entry point,
/// and a profile with several mailboxes has several independent rule collections.
/// </para>
///
/// <para>
/// <b>A release rule specific to this object graph.</b> Everything reachable below a
/// <c>Rules</c> collection - the <c>Rule</c> objects, their <c>Conditions</c> and <c>Actions</c>
/// collections, and every clause slot inside those - is a cached child that Outlook owns and hands
/// back by the same pointer on every access. <c>Rules.Item(3)</c> and a later <c>Rules.Item(3)</c>
/// are one object, and so are <c>Conditions.Subject</c> and the <c>Conditions[n]</c> that reports
/// the subject slot. Final-releasing any of them drops the refcount on an object Outlook is still
/// handing out, so the next access is a use-after-free that takes the host process down rather than
/// throwing something catchable. They are therefore released with
/// <c>ReleaseSharedComObject</c> - a plain decrement - for the same reason #19 requires it for the
/// shared <c>Application</c>. Only the objects this class genuinely owns the lifetime of (the
/// <c>Rules</c> collection itself, the <c>Store</c>, and a destination folder it navigated to) are
/// final-released.
/// </para>
/// </summary>
public class RuleCommands : IRuleCommands
{
    /// <summary>
    /// Rules only ever run against arriving mail in this surface. Send rules exist in Outlook and
    /// are enumerated correctly by <see cref="List"/>, but creating one is a different mental model
    /// (it fires on messages the user is sending) and is deliberately not offered.
    /// </summary>
    private const Outlook.OlRuleType CreatedRuleType = Outlook.OlRuleType.olRuleReceive;

    private static readonly Dictionary<string, Outlook.OlDefaultFolders> FolderAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["inbox"] = Outlook.OlDefaultFolders.olFolderInbox,
            ["drafts"] = Outlook.OlDefaultFolders.olFolderDrafts,
            ["sent"] = Outlook.OlDefaultFolders.olFolderSentMail,
            ["sentitems"] = Outlook.OlDefaultFolders.olFolderSentMail,
            ["outbox"] = Outlook.OlDefaultFolders.olFolderOutbox,
            ["deleted"] = Outlook.OlDefaultFolders.olFolderDeletedItems,
            ["deleteditems"] = Outlook.OlDefaultFolders.olFolderDeletedItems,
            ["junk"] = Outlook.OlDefaultFolders.olFolderJunk,
            ["archive"] = Outlook.OlDefaultFolders.olFolderInbox
        };

    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailRuleListResult List(bool includeDetail = false, string? storeId = null)
    {
        return OutlookInteropRunner.Execute(
            "OutlookRuleList",
            (application, session) =>
            {
                Outlook.Store? store = null;
                Outlook.Rules? rules = null;

                try
                {
                    store = ResolveStore(session, storeId, out string? storeError);
                    if (store == null)
                    {
                        return new MailRuleListResult { Success = false, ErrorMessage = storeError };
                    }

                    var result = new MailRuleListResult
                    {
                        Success = true,
                        StoreDisplayName = SafeGet(() => store.DisplayName)
                    };

                    rules = store.GetRules();
                    int count = rules.Count;

                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Rule? rule = null;

                        try
                        {
                            rule = rules[index];

                            var info = DescribeRule(rule, includeDetail);
                            if (info != null)
                            {
                                result.Rules.Add(info);
                            }
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseSharedComObject(ref rule);
                        }
                    }

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref rules);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            },
            ex => new MailRuleListResult
            {
                Success = false,
                ErrorMessage = $"Failed to read the Outlook rule list: {ex.Message}"
            });
    }

    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailRuleMutationResult Create(
        string name,
        string? fromAddress = null,
        string? subjectContains = null,
        string? moveToFolder = null,
        string? assignCategories = null,
        bool deleteMessage = false,
        bool stopProcessingRules = false,
        bool enabled = true,
        string? storeId = null)
    {
        string? trimmedName = NullIfBlank(name);
        if (trimmedName == null)
        {
            return Refuse("A rule name is required. Outlook shows it in the rule list and this surface uses it to address the rule.");
        }

        trimmedName = trimmedName.Trim();

        bool hasCondition = NullIfBlank(fromAddress) != null || NullIfBlank(subjectContains) != null;
        if (!hasCondition)
        {
            return Refuse(
                $"Rule '{trimmedName}' would have no conditions, so it would match every message that arrives. "
                + "Give it at least one of fromAddress or subjectContains.");
        }

        bool hasAction = NullIfBlank(moveToFolder) != null
            || NullIfBlank(assignCategories) != null
            || deleteMessage
            || stopProcessingRules;

        if (!hasAction)
        {
            return Refuse(
                $"Rule '{trimmedName}' would have no actions, so it would do nothing to the mail it matched. "
                + "Give it at least one of moveToFolder, assignCategories, deleteMessage or stopProcessingRules.");
        }

        if (deleteMessage && NullIfBlank(moveToFolder) != null)
        {
            return Refuse(DeleteAndMoveConflict);
        }

        return OutlookInteropRunner.Execute(
            "OutlookRuleCreate",
            (application, session) =>
            {
                Outlook.Store? store = null;
                Outlook.Rules? rules = null;
                Outlook.Rule? rule = null;
                Outlook.MAPIFolder? destination = null;
                Outlook.Explorer? explorer = null;

                try
                {
                    store = ResolveStore(session, storeId, out string? storeError);
                    if (store == null)
                    {
                        return Refuse(storeError!);
                    }

                    rules = store.GetRules();

                    if (FindRuleIndexes(rules, trimmedName).Count > 0)
                    {
                        return Refuse(
                            $"This mailbox already has a rule named '{trimmedName}'. Outlook allows duplicate rule "
                            + "names, but then no later update or delete can say which one it means, so a second "
                            + "rule with the same name is refused. Pick another name, or use update to change the "
                            + "existing rule.");
                    }

                    if (NullIfBlank(moveToFolder) != null)
                    {
                        destination = OutlookInteropRunner.ResolveFolder(
                            application, session, moveToFolder, FolderAliases, ref explorer);

                        if (destination == null)
                        {
                            return Refuse(
                                $"No folder matches '{moveToFolder}'. A move rule cannot be created against a folder "
                                + "that does not exist yet - create it first with folder create.");
                        }
                    }

                    // Everything that can be refused has been refused. Only now is the collection
                    // touched, so a refusal above can never leave a half-built rule behind.
                    rule = rules.Create(trimmedName, CreatedRuleType);
                    rule.Enabled = enabled;

                    ApplyConditions(rule, fromAddress, subjectContains);
                    ApplyActions(rule, destination, assignCategories, deleteMessage, stopProcessingRules);

                    Save(rules);

                    return Describe(rules, store, trimmedName, "created");
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref destination);
                    OutlookInteropRunner.ReleaseSharedComObject(ref rule);
                    OutlookInteropRunner.ReleaseComObject(ref rules);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            },
            ex => FailedSave($"create rule '{trimmedName}'", ex));
    }

    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailRuleMutationResult Update(
        string name,
        string? fromAddress = null,
        string? subjectContains = null,
        string? moveToFolder = null,
        string? assignCategories = null,
        bool? deleteMessage = null,
        bool? stopProcessingRules = null,
        string? newName = null,
        string? storeId = null)
    {
        string? trimmedName = NullIfBlank(name)?.Trim();
        if (trimmedName == null)
        {
            return Refuse("The name of the rule to update is required.");
        }

        string? trimmedNewName = NullIfBlank(newName)?.Trim();

        if (deleteMessage == true && NullIfBlank(moveToFolder) != null)
        {
            return Refuse(DeleteAndMoveConflict);
        }

        return OutlookInteropRunner.Execute(
            "OutlookRuleUpdate",
            (application, session) =>
            {
                Outlook.Store? store = null;
                Outlook.Rules? rules = null;
                Outlook.Rule? rule = null;
                Outlook.MAPIFolder? destination = null;
                Outlook.Explorer? explorer = null;

                try
                {
                    store = ResolveStore(session, storeId, out string? storeError);
                    if (store == null)
                    {
                        return Refuse(storeError!);
                    }

                    rules = store.GetRules();

                    if (!TryResolveSingleRule(rules, trimmedName, out int index, out string? resolveError))
                    {
                        return Refuse(resolveError!);
                    }

                    if (trimmedNewName != null
                        && !string.Equals(trimmedNewName, trimmedName, StringComparison.Ordinal)
                        && FindRuleIndexes(rules, trimmedNewName).Count > 0)
                    {
                        return Refuse(
                            $"This mailbox already has a rule named '{trimmedNewName}', and two rules sharing a name "
                            + "cannot afterwards be told apart by name.");
                    }

                    if (NullIfBlank(moveToFolder) != null)
                    {
                        destination = OutlookInteropRunner.ResolveFolder(
                            application, session, moveToFolder, FolderAliases, ref explorer);

                        if (destination == null)
                        {
                            return Refuse($"No folder matches '{moveToFolder}'.");
                        }
                    }

                    rule = rules[index];

                    // Read the rule's current clauses once, before anything is written. Two reasons,
                    // and the second is not optional: the projection below needs the clauses this
                    // patch does not mention, and re-reading rule.Conditions after the patch would
                    // ask Outlook for a child object whose RCW this call has already final-released.
                    // Outlook hands back the same underlying pointer, so that is a use-after-free
                    // that takes the host process down rather than throwing.
                    var current = DescribeRule(rule, includeDetail: true);

                    if (current == null)
                    {
                        return Refuse($"Rule '{trimmedName}' could not be read back from the mailbox.");
                    }

                    // The same invariant create enforces, applied to the rule the patch would
                    // produce: a patch must not be able to strip a rule down to matching every
                    // message, or to doing nothing at all.
                    if (!ProjectClause(current.Conditions, ConditionSlots,
                            (SenderAddressCondition, fromAddress), (SubjectCondition, subjectContains)))
                    {
                        return Refuse(
                            $"That update would leave rule '{trimmedName}' with no conditions, so it would match "
                            + "every message that arrives. Nothing was changed.");
                    }

                    if (!ProjectClause(current.Actions, ActionSlots,
                            (MoveToFolderAction, moveToFolder),
                            (AssignToCategoryAction, assignCategories),
                            (DeleteAction, AsClauseValue(deleteMessage)),
                            (StopAction, AsClauseValue(stopProcessingRules))))
                    {
                        return Refuse(
                            $"That update would leave rule '{trimmedName}' with no actions, so it would do nothing "
                            + "to the mail it matched. Nothing was changed.");
                    }

                    ApplyConditionPatch(rule, fromAddress, subjectContains);
                    ApplyActionPatch(rule, moveToFolder, destination, assignCategories, deleteMessage, stopProcessingRules);

                    if (trimmedNewName != null)
                    {
                        rule.Name = trimmedNewName;
                    }

                    Save(rules);

                    return Describe(rules, store, trimmedNewName ?? trimmedName, "updated");
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref destination);
                    OutlookInteropRunner.ReleaseSharedComObject(ref rule);
                    OutlookInteropRunner.ReleaseComObject(ref rules);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            },
            ex => FailedSave($"update rule '{trimmedName}'", ex));
    }

    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailRuleMutationResult SetEnabled(string name, bool enabled = true, string? storeId = null)
    {
        string? trimmedName = NullIfBlank(name)?.Trim();
        if (trimmedName == null)
        {
            return Refuse("The name of the rule to switch on or off is required.");
        }

        return OutlookInteropRunner.Execute(
            "OutlookRuleSetEnabled",
            (application, session) =>
            {
                Outlook.Store? store = null;
                Outlook.Rules? rules = null;
                Outlook.Rule? rule = null;

                try
                {
                    store = ResolveStore(session, storeId, out string? storeError);
                    if (store == null)
                    {
                        return Refuse(storeError!);
                    }

                    rules = store.GetRules();

                    if (!TryResolveSingleRule(rules, trimmedName, out int index, out string? resolveError))
                    {
                        return Refuse(resolveError!);
                    }

                    rule = rules[index];
                    rule.Enabled = enabled;

                    Save(rules);

                    return Describe(rules, store, trimmedName, enabled ? "enabled" : "disabled");
                }
                finally
                {
                    OutlookInteropRunner.ReleaseSharedComObject(ref rule);
                    OutlookInteropRunner.ReleaseComObject(ref rules);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            },
            ex => FailedSave($"{(enabled ? "enable" : "disable")} rule '{trimmedName}'", ex));
    }

    /// <inheritdoc/>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailRuleMutationResult Delete(string name, string? storeId = null)
    {
        string? trimmedName = NullIfBlank(name)?.Trim();
        if (trimmedName == null)
        {
            return Refuse("The name of the rule to delete is required.");
        }

        return OutlookInteropRunner.Execute(
            "OutlookRuleDelete",
            (application, session) =>
            {
                Outlook.Store? store = null;
                Outlook.Rules? rules = null;

                try
                {
                    store = ResolveStore(session, storeId, out string? storeError);
                    if (store == null)
                    {
                        return Refuse(storeError!);
                    }

                    rules = store.GetRules();

                    if (!TryResolveSingleRule(rules, trimmedName, out int index, out string? resolveError))
                    {
                        return Refuse(resolveError!);
                    }

                    rules.Remove(index);

                    Save(rules);

                    return new MailRuleMutationResult
                    {
                        Success = true,
                        Action = "deleted",
                        Name = trimmedName,
                        StoreDisplayName = SafeGet(() => store.DisplayName),
                        RuleCount = SafeGetInt(() => rules.Count)
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref rules);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            },
            ex => FailedSave($"delete rule '{trimmedName}'", ex));
    }

    // ── Persistence ──────────────────────────────────────────

    /// <summary>
    /// Commits the store's whole rule collection.
    ///
    /// <para>
    /// <c>ShowProgress</c> is false deliberately. The progress dialog is modal and belongs to
    /// Outlook, so on a slow Exchange connection it would sit on the user's desktop waiting for an
    /// automated caller that cannot dismiss it, and cancelling it makes the save fail anyway.
    /// Suppressing it does not remove every prompt - Outlook may still raise its own modal warning
    /// about rule compatibility, which nothing here can answer; that case surfaces as the
    /// dispatcher's operation timeout rather than as a hang.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void Save(Outlook.Rules rules) => rules.Save(false);

    /// <summary>
    /// The failure message for a mutation. Says explicitly that nothing was written, because that is
    /// the one thing a caller needs to know and cannot infer: an unsaved <c>Rules</c> collection is
    /// discarded whole, so a failed save leaves the mailbox exactly as it was rather than partly
    /// changed.
    /// </summary>
    private static MailRuleMutationResult FailedSave(string what, Exception ex) =>
        new()
        {
            Success = false,
            ErrorMessage =
                $"Failed to {what}: {ex.Message} No change was written to the mailbox - Outlook commits a store's "
                + "rules as one collection, so a failed save leaves every rule as it was. Common causes are the "
                + "Rules and Alerts wizard being open on this mailbox, the Exchange rules quota being full, or "
                + "another rule in the mailbox being malformed."
        };

    private static MailRuleMutationResult Refuse(string message) =>
        new() { Success = false, ErrorMessage = message };

    /// <summary>
    /// Outlook does not implement "delete it" as its own action: it rewrites it into a move to
    /// Deleted Items, and there is exactly one move slot per rule. Asking for both silently discards
    /// one of them, so both together are refused rather than half-honoured.
    /// </summary>
    private const string DeleteAndMoveConflict =
        "deleteMessage and moveToFolder cannot both be set. Outlook stores 'delete it' as a move to Deleted Items "
        + "and a rule has only one move destination, so asking for both would silently drop one. Use moveToFolder "
        + "on its own, or deleteMessage on its own.";

    // ── Clause projection ────────────────────────────────────

    // The clause names as they come back from list, which is how the projection below lines a
    // patch's arguments up against the clauses a rule already has.
    private const string SenderAddressCondition = "senderAddress";
    private const string SubjectCondition = "subject";
    private const string MoveToFolderAction = "moveToFolder";
    private const string AssignToCategoryAction = "assignToCategory";
    private const string DeleteAction = "delete";
    private const string StopAction = "stop";

    private static readonly string[] ConditionSlots = [SenderAddressCondition, SubjectCondition];

    private static readonly string[] ActionSlots =
        [MoveToFolderAction, AssignToCategoryAction, DeleteAction, StopAction];

    /// <summary>
    /// Whether the rule would still have at least one clause of this kind after the patch.
    ///
    /// <para>
    /// Clauses this surface cannot write - a <c>from</c> condition, a <c>body</c> condition, a
    /// forward action - are counted as they stand. Refusing an update because the rule's only
    /// condition happens to be one this tool cannot express would block a caller from touching
    /// rules the user built in Outlook, which is most of them.
    /// </para>
    /// </summary>
    private static bool ProjectClause(
        List<string> existing,
        string[] writableSlots,
        params (string Slot, string? Value)[] patch)
    {
        bool untouched = existing.Any(clause => !writableSlots.Contains(clause, StringComparer.Ordinal));
        if (untouched)
        {
            return true;
        }

        foreach (var (slot, value) in patch)
        {
            bool present = value == null
                ? existing.Contains(slot, StringComparer.Ordinal)
                : NullIfBlank(value) != null;

            if (present)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Renders a tri-state action switch as the same "null means unchanged, blank means off" shape
    /// the string clauses use, so <see cref="ProjectClause"/> has one rule rather than two.
    /// </summary>
    private static string? AsClauseValue(bool? enabled) =>
        enabled switch
        {
            null => null,
            true => "on",
            false => string.Empty
        };

    /// <summary>
    /// Reads back what Outlook actually stored, rather than echoing what was asked for. A mutation
    /// that reported the requested state would be unfalsifiable.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static MailRuleMutationResult Describe(
        Outlook.Rules rules,
        Outlook.Store store,
        string name,
        string action)
    {
        Outlook.Rule? saved = null;

        try
        {
            var indexes = FindRuleIndexes(rules, name);
            if (indexes.Count == 1)
            {
                saved = rules[indexes[0]];
            }

            return new MailRuleMutationResult
            {
                Success = true,
                Action = action,
                Name = name,
                Enabled = saved != null ? SafeGetBool(() => saved.Enabled) : null,
                ExecutionOrder = saved != null ? SafeGetInt(() => saved.ExecutionOrder) : null,
                StoreDisplayName = SafeGet(() => store.DisplayName),
                RuleCount = SafeGetInt(() => rules.Count)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref saved);
        }
    }

    // ── Clause writing ───────────────────────────────────────

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ApplyConditions(Outlook.Rule rule, string? fromAddress, string? subjectContains)
    {
        Outlook.RuleConditions? conditions = null;

        try
        {
            conditions = rule.Conditions;
            SetAddressCondition(conditions.SenderAddress, fromAddress);
            SetTextCondition(conditions.Subject, subjectContains);
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref conditions);
        }
    }

    /// <summary>
    /// PATCH semantics: null leaves a condition exactly as it was, an empty string clears it.
    /// Without the distinction a condition could be added but never removed.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ApplyConditionPatch(Outlook.Rule rule, string? fromAddress, string? subjectContains)
    {
        Outlook.RuleConditions? conditions = null;

        try
        {
            conditions = rule.Conditions;

            if (fromAddress != null)
            {
                SetAddressCondition(conditions.SenderAddress, NullIfBlank(fromAddress));
            }

            if (subjectContains != null)
            {
                SetTextCondition(conditions.Subject, NullIfBlank(subjectContains));
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref conditions);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ApplyActions(
        Outlook.Rule rule,
        Outlook.MAPIFolder? destination,
        string? assignCategories,
        bool deleteMessage,
        bool stopProcessingRules)
    {
        Outlook.RuleActions? actions = null;

        try
        {
            actions = rule.Actions;
            SetMoveAction(actions.MoveToFolder, destination);
            SetCategoryAction(actions.AssignToCategory, SplitCategories(assignCategories));
            SetFlagAction(actions.Delete, deleteMessage);
            SetFlagAction(actions.Stop, stopProcessingRules);
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref actions);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ApplyActionPatch(
        Outlook.Rule rule,
        string? moveToFolder,
        Outlook.MAPIFolder? destination,
        string? assignCategories,
        bool? deleteMessage,
        bool? stopProcessingRules)
    {
        Outlook.RuleActions? actions = null;

        try
        {
            actions = rule.Actions;

            if (moveToFolder != null)
            {
                SetMoveAction(actions.MoveToFolder, destination);
            }

            if (assignCategories != null)
            {
                SetCategoryAction(actions.AssignToCategory, SplitCategories(assignCategories));
            }

            if (deleteMessage.HasValue)
            {
                SetFlagAction(actions.Delete, deleteMessage.Value);
            }

            if (stopProcessingRules.HasValue)
            {
                SetFlagAction(actions.Stop, stopProcessingRules.Value);
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref actions);
        }
    }

    /// <summary>
    /// Sets the sender-address condition.
    ///
    /// <para>
    /// Deliberately <c>SenderAddress</c> and not <c>From</c>. Outlook's <c>From</c> condition holds
    /// address-book entries and needs <c>Recipients.ResolveAll</c> to be usable, which is a
    /// protected member: an out-of-process caller touching it raises the Object Model Guard prompt,
    /// which cannot be answered programmatically, so the write would hang and then time out.
    /// <c>SenderAddress</c> is a plain substring match on the SMTP address, needs no resolution, and
    /// is what a caller almost always means by "mail from this person" anyway.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void SetAddressCondition(Outlook.AddressRuleCondition condition, string? address)
    {
        try
        {
            if (address == null)
            {
                condition.Enabled = false;
                return;
            }

            condition.Address = new[] { address };
            condition.Enabled = true;
        }
        finally
        {
            var releasable = condition;
            OutlookInteropRunner.ReleaseSharedComObject(ref releasable);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void SetTextCondition(Outlook.TextRuleCondition condition, string? text)
    {
        try
        {
            if (text == null)
            {
                condition.Enabled = false;
                return;
            }

            condition.Text = new[] { text };
            condition.Enabled = true;
        }
        finally
        {
            var releasable = condition;
            OutlookInteropRunner.ReleaseSharedComObject(ref releasable);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void SetMoveAction(Outlook.MoveOrCopyRuleAction action, Outlook.MAPIFolder? destination)
    {
        try
        {
            if (destination == null)
            {
                action.Enabled = false;
                return;
            }

            action.Folder = destination;
            action.Enabled = true;
        }
        finally
        {
            var releasable = action;
            OutlookInteropRunner.ReleaseSharedComObject(ref releasable);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void SetCategoryAction(Outlook.AssignToCategoryRuleAction action, string[]? categories)
    {
        try
        {
            if (categories == null || categories.Length == 0)
            {
                action.Enabled = false;
                return;
            }

            action.Categories = categories;
            action.Enabled = true;
        }
        finally
        {
            var releasable = action;
            OutlookInteropRunner.ReleaseSharedComObject(ref releasable);
        }
    }

    /// <summary>
    /// For the actions that carry no payload - delete and stop-processing - where being switched on
    /// is the whole of the instruction.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void SetFlagAction(Outlook.RuleAction action, bool enabled)
    {
        try
        {
            action.Enabled = enabled;
        }
        finally
        {
            var releasable = action;
            OutlookInteropRunner.ReleaseSharedComObject(ref releasable);
        }
    }

    private static string[]? SplitCategories(string? categories)
    {
        if (NullIfBlank(categories) == null)
        {
            return null;
        }

        return categories!
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    }

    // ── Rule lookup ──────────────────────────────────────────

    /// <summary>
    /// The indexes of every rule with this exact name. A list rather than a single index because
    /// Outlook permits duplicates, and a lookup that quietly returned the first match would let an
    /// update or a delete hit a rule the caller never meant.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static List<int> FindRuleIndexes(Outlook.Rules rules, string name)
    {
        var matches = new List<int>();
        int count = SafeGetInt(() => rules.Count);

        for (int index = 1; index <= count; index++)
        {
            Outlook.Rule? rule = null;

            try
            {
                rule = rules[index];
                if (string.Equals(SafeGet(() => rule.Name), name, StringComparison.Ordinal))
                {
                    matches.Add(index);
                }
            }
            finally
            {
                OutlookInteropRunner.ReleaseSharedComObject(ref rule);
            }
        }

        return matches;
    }

    private static bool TryResolveSingleRule(
        Outlook.Rules rules,
        string name,
        out int index,
        out string? error)
    {
        var matches = FindRuleIndexes(rules, name);

        if (matches.Count == 0)
        {
            index = 0;
            error =
                $"This mailbox has no rule named '{name}'. Rule names are matched exactly, including case and "
                + "spacing; use list to see the names as Outlook holds them.";
            return false;
        }

        if (matches.Count > 1)
        {
            index = 0;
            error =
                $"This mailbox has {matches.Count} rules named '{name}', so there is no way to tell which one was "
                + "meant. Rename them in Outlook before changing either.";
            return false;
        }

        index = matches[0];
        error = null;
        return true;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Store? ResolveStore(Outlook.NameSpace session, string? storeId, out string? error)
    {
        error = null;

        if (NullIfBlank(storeId) == null)
        {
            return session.DefaultStore;
        }

        Outlook.Stores? stores = null;

        try
        {
            stores = session.Stores;
            int count = SafeGetInt(() => stores.Count);

            for (int index = 1; index <= count; index++)
            {
                Outlook.Store? store = null;
                bool keep = false;

                try
                {
                    store = stores[index];
                    keep = string.Equals(SafeGet(() => store.StoreID), storeId, StringComparison.Ordinal);

                    if (keep)
                    {
                        return store;
                    }
                }
                finally
                {
                    if (!keep)
                    {
                        OutlookInteropRunner.ReleaseComObject(ref store);
                    }
                }
            }

            // Falling back to the default store here would rewrite the rules of a mailbox the caller
            // did not name, under success: true. See folder list-default, which refuses for the same
            // reason.
            error =
                $"No store in this Outlook profile has the id '{storeId}'. Use folder list-stores to discover the "
                + "available store ids.";
            return null;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref stores);
        }
    }

    // ── Clause reading ───────────────────────────────────────

    /// <summary>
    /// Renders one rule. Shared by <see cref="List"/> and by the update guard, so the invariant the
    /// guard enforces is checked against exactly what a caller would see.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static MailRuleInfo? DescribeRule(Outlook.Rule rule, bool includeDetail)
    {
        string? name = SafeGet(() => rule.Name);
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        var info = new MailRuleInfo
        {
            Name = name,
            Enabled = SafeGetBool(() => rule.Enabled),
            ExecutionOrder = SafeGetInt(() => rule.ExecutionOrder),
            RuleType = SafeGetInt(() => (int)rule.RuleType) == (int)Outlook.OlRuleType.olRuleSend
                ? "send"
                : "receive",
            IsLocalRule = SafeGetBool(() => rule.IsLocalRule)
        };

        if (includeDetail)
        {
            ReadConditions(rule, info);
            ReadActions(rule, info);
        }

        return info;
    }

    /// <summary>
    /// Walks a rule's fixed condition slots and keeps only the ones switched on, together with the
    /// values they match against.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadConditions(Outlook.Rule rule, MailRuleInfo info)
    {
        Outlook.RuleConditions? conditions = null;

        try
        {
            conditions = rule.Conditions;
            int count = SafeGetInt(() => conditions.Count);

            for (int index = 1; index <= count; index++)
            {
                Outlook.RuleCondition? condition = null;

                try
                {
                    condition = conditions[index];

                    if (!SafeGetBool(() => condition.Enabled))
                    {
                        continue;
                    }

                    int conditionType = SafeGetInt(() => (int)condition.ConditionType);
                    string? conditionName = StripEnumPrefix(
                        Enum.GetName(typeof(Outlook.OlRuleConditionType), conditionType),
                        "olCondition");

                    if (conditionName == null)
                    {
                        continue;
                    }

                    info.Conditions.Add(conditionName);

                    switch (condition)
                    {
                        case Outlook.ToOrFromRuleCondition addressCondition
                            when conditionType == (int)Outlook.OlRuleConditionType.olConditionFrom:
                            ReadRuleRecipients(addressCondition, info.FromAddresses);
                            break;

                        case Outlook.AddressRuleCondition senderCondition
                            when conditionType == (int)Outlook.OlRuleConditionType.olConditionSenderAddress:
                            AddAll(info.SenderAddresses, ReadStringArray(() => senderCondition.Address));
                            break;

                        case Outlook.TextRuleCondition textCondition
                            when conditionType == (int)Outlook.OlRuleConditionType.olConditionSubject:
                            AddAll(info.SubjectTerms, ReadStringArray(() => textCondition.Text));
                            break;
                    }
                }
                finally
                {
                    OutlookInteropRunner.ReleaseSharedComObject(ref condition);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref conditions);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadActions(Outlook.Rule rule, MailRuleInfo info)
    {
        Outlook.RuleActions? actions = null;

        try
        {
            actions = rule.Actions;
            int count = SafeGetInt(() => actions.Count);

            for (int index = 1; index <= count; index++)
            {
                Outlook.RuleAction? action = null;

                try
                {
                    action = actions[index];

                    if (!SafeGetBool(() => action.Enabled))
                    {
                        continue;
                    }

                    int actionType = SafeGetInt(() => (int)action.ActionType);
                    string? actionName = StripEnumPrefix(
                        Enum.GetName(typeof(Outlook.OlRuleActionType), actionType),
                        "olRuleAction");

                    if (actionName == null)
                    {
                        continue;
                    }

                    info.Actions.Add(actionName);

                    switch (action)
                    {
                        case Outlook.MoveOrCopyRuleAction moveAction
                            when actionType == (int)Outlook.OlRuleActionType.olRuleActionMoveToFolder:
                            ReadMoveDestination(moveAction, info);
                            break;

                        case Outlook.AssignToCategoryRuleAction categoryAction:
                            AddAll(info.AssignCategories, ReadStringArray(() => categoryAction.Categories));
                            break;
                    }
                }
                finally
                {
                    OutlookInteropRunner.ReleaseSharedComObject(ref action);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref actions);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadMoveDestination(Outlook.MoveOrCopyRuleAction moveAction, MailRuleInfo info)
    {
        Outlook.MAPIFolder? folder = null;

        try
        {
            folder = SafeGetComObject(() => moveAction.Folder);
            if (folder != null)
            {
                info.MoveToFolderPath ??= SafeGet(() => OutlookInteropRunner.GetFolderPath(folder));
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref folder);
        }
    }

    /// <summary>
    /// Rule recipients are stored unresolved, so Outlook leaves <c>Address</c> blank and puts the
    /// address in <c>Name</c>. Reading only <c>Address</c> reports a from-rule as matching nobody.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ReadRuleRecipients(Outlook.ToOrFromRuleCondition condition, List<string> addresses)
    {
        Outlook.Recipients? recipients = null;

        try
        {
            recipients = condition.Recipients;
            int count = SafeGetInt(() => recipients.Count);

            for (int index = 1; index <= count; index++)
            {
                Outlook.Recipient? recipient = null;

                try
                {
                    recipient = recipients[index];

                    string? address = NullIfBlank(SafeGet(() => recipient.Address))
                        ?? NullIfBlank(SafeGet(() => recipient.Name));

                    if (address != null && !addresses.Contains(address, StringComparer.OrdinalIgnoreCase))
                    {
                        addresses.Add(address);
                    }
                }
                finally
                {
                    OutlookInteropRunner.ReleaseSharedComObject(ref recipient);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseSharedComObject(ref recipients);
        }
    }

    /// <summary>
    /// Rule clause values come back as a COM SAFEARRAY, which the runtime surfaces as
    /// <c>string[]</c> in the common case and as <c>object[]</c> when the variant type is not
    /// uniform. Handling only the first shape reports a populated clause as empty.
    /// </summary>
    private static List<string> ReadStringArray(Func<object?> getter)
    {
        object? raw;

        try
        {
            raw = getter();
        }
        catch
        {
            return [];
        }

        return raw switch
        {
            string single => [single],
            string[] strings => [.. strings.Where(s => !string.IsNullOrWhiteSpace(s))],
            System.Collections.IEnumerable values =>
                [.. values.Cast<object?>().Select(v => v?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s))!],
            _ => []
        };
    }

    private static void AddAll(List<string> target, List<string> values)
    {
        foreach (string value in values)
        {
            if (!target.Contains(value, StringComparer.Ordinal))
            {
                target.Add(value);
            }
        }
    }

    // ── Property-read helpers ────────────────────────────────

    private static string? SafeGet(Func<string?> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static bool SafeGetBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return false;
        }
    }

    private static int SafeGetInt(Func<int> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return 0;
        }
    }

    private static T? SafeGetComObject<T>(Func<T?> getter)
        where T : class
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static string? NullIfBlank(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private static string? StripEnumPrefix(string? enumName, string prefix)
    {
        if (string.IsNullOrEmpty(enumName) || !enumName.StartsWith(prefix, StringComparison.Ordinal))
        {
            return null;
        }

        string trimmed = enumName[prefix.Length..];
        return trimmed.Length == 0
            ? null
            : char.ToLowerInvariant(trimmed[0]) + trimmed[1..];
    }
}
