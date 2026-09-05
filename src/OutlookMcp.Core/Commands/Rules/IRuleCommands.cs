using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Rules;

/// <summary>
/// Outlook inbox rule operations.
/// </summary>
///
/// <remarks>
/// <para>
/// <b>Why this is its own category rather than more actions on <c>mail</c>.</b> Three reasons, in
/// order of weight. Every <c>mail</c> action addresses a <em>message</em> by <c>entryId</c> and
/// <c>storeId</c>; every action here addresses a <em>rule</em> by name within a store's rule
/// collection. Putting two identity models in one tool is how a caller ends up passing an entry id
/// where a rule name belongs. Second, the blast radius differs in kind: <c>mail delete</c> affects
/// one message that is recoverable from Deleted Items, while a wrong rule silently redirects or
/// destroys mail that has not arrived yet, for as long as nobody notices. That deserves a tool
/// description leading with the warning rather than a clause buried in a four-thousand-character
/// one. Third, <c>mail</c>'s description is already at the length where an LLM stops reading
/// carefully, and rule semantics need several hundred more characters that would degrade every
/// other mail action to add.
/// </para>
///
/// <para>
/// <c>mail list-rules</c> is kept and now delegates here, so there is one implementation and no
/// behaviour change for anything already calling it. It earns its place in the <c>mail</c> tool
/// because "why is nothing arriving from this sender?" is a mail question whose answer is a rule.
/// </para>
///
/// <para>
/// The summary above is deliberately one short line: the CLI generator turns it into the command's
/// one-line description in <c>outlookcli --help</c>, where a rationale paragraph would swamp the
/// command list. The reasoning belongs here, and the caller-facing guidance belongs in the
/// <c>McpTool</c> description below.
/// </para>
/// </remarks>
[ServiceCategory("rule")]
[McpTool("rule", Title = "Outlook Rule Operations", Destructive = true, Category = "rule",
    Description = "Read and change the inbox rules that decide what happens to mail before anyone reads it. "
    + "Treat every write here as higher-risk than deleting a message. A message deleted in error sits in Deleted "
    + "Items; a rule created in error silently moves or destroys mail that has not arrived yet, keeps doing it, and "
    + "is typically noticed days later. Prefer list, and confirm with the user before create, update, set-enabled "
    + "or delete. "
    + "Use list to enumerate a mailbox's rules. Pass includeDetail for each rule's conditions, actions, subject "
    + "terms, sender addresses and move-to destination; it is off by default because gathering it is roughly forty "
    + "times the work. "
    + "Rules are per-store, never global. Every action defaults to the profile's default delivery store; pass a "
    + "storeId from folder list-stores to reach another mailbox. An unknown storeId is refused rather than "
    + "silently falling back to the default store. "
    + "Rules are addressed by name. Outlook permits two rules to share a name, so create refuses a name already in "
    + "use, and update, set-enabled and delete refuse a name that matches no rule or more than one. "
    + "create requires at least one condition and at least one action. A rule with no conditions matches every "
    + "message that arrives, and one with no actions does nothing at all; Outlook accepts both and this refuses "
    + "them. New rules are enabled unless enabled=false is passed, and Outlook inserts them at the TOP of the "
    + "evaluation order, so a new rule runs before every rule the mailbox already had - if it stops processing, "
    + "nothing else will run. "
    + "Supported conditions are fromAddress (a substring of the sender's SMTP address) and subjectContains (a "
    + "substring of the subject). Supported actions are moveToFolder, assignCategories, deleteMessage and "
    + "stopProcessingRules. Each condition takes one term - a rule matching several terms can be read back by list "
    + "but cannot be written here. "
    + "deleteMessage does not read back as a delete. Outlook has no delete action: it rewrites 'delete it' into a "
    + "move to Deleted Items plus stop-processing, so list afterwards reports moveToFolder with a Deleted Items "
    + "destination. That is the rule working correctly, not a different rule. For the same reason deleteMessage "
    + "and moveToFolder cannot both be set - a rule has one move destination - and setting deleteMessage false "
    + "later does not necessarily remove the move Outlook created from it; clear it with moveToFolder set to an "
    + "empty string. "
    + "There is deliberately no mark-as-read action: Outlook's rule object model has none, so no caller of any "
    + "kind can create one, and only the Rules and Alerts wizard inside Outlook can. There is also deliberately no "
    + "forward, redirect or CC action, because those send mail on the user's behalf, unattended, indefinitely. "
    + "update leaves any clause it is not given alone. Pass an empty string to clear a condition, or false to "
    + "switch an action off. "
    + "Writes are not per-rule: Outlook commits a store's whole rule collection at once, so every write here "
    + "rewrites all of the mailbox's rules and the response reports ruleCount so the caller can check the total is "
    + "what they expected. That save is also the only step that persists anything, and it is the step that fails - "
    + "on an Exchange rules quota, or because the user has the Rules and Alerts wizard open, or because some other "
    + "rule in the mailbox is malformed. When it fails, nothing was written at all, not even partially, and the "
    + "error says so.")]
public interface IRuleCommands
{
    /// <summary>
    /// Enumerates a mailbox's rules.
    /// </summary>
    /// <param name="includeDetail">Gather each rule's conditions, actions, subject terms, sender addresses and move-to destination. Off by default: Outlook's condition and action collections have a fixed length covering every clause it supports, so detail means walking roughly 59 slots per rule.</param>
    /// <param name="storeId">The mailbox to read, from <c>folder list-stores</c>. Defaults to the profile's default delivery store.</param>
    [ServiceAction("list", Destructive = false)]
    MailRuleListResult List(bool includeDetail = false, string? storeId = null);

    /// <summary>
    /// Creates a rule and commits it.
    ///
    /// <para>
    /// Requires at least one condition and at least one action. New rules are enabled by default,
    /// matching what the Rules and Alerts wizard does, so a caller who wants to stage a rule without
    /// it acting on mail must pass <c>enabled: false</c> deliberately.
    /// </para>
    ///
    /// <para>
    /// Outlook inserts a new rule at the <em>top</em> of the evaluation order, not the bottom. On a
    /// mailbox with an established rule set that is the opposite of what most callers assume, and it
    /// matters: a new rule runs before all of them, and one that also stops processing prevents
    /// every existing rule from running at all.
    /// </para>
    /// </summary>
    /// <param name="name">The rule's name, as it will appear in Outlook. Must not already be in use in this store.</param>
    /// <param name="fromAddress">Match when the sender's SMTP address contains this. A substring match on the address itself, so no address-book lookup is involved.</param>
    /// <param name="subjectContains">Match when the subject contains this.</param>
    /// <param name="moveToFolder">Move matching mail to this folder - a default folder role such as 'inbox' or a full folder path. The folder must already exist.</param>
    /// <param name="assignCategories">Stamp matching mail with these categories, comma-separated. Use <c>mail list-categories</c> to discover which names exist.</param>
    /// <param name="deleteMessage">Move matching mail to Deleted Items. Never a permanent delete. Outlook stores this as a move plus stop-processing rather than as a delete action, so <c>list</c> reports it as <c>moveToFolder</c>. Cannot be combined with <paramref name="moveToFolder"/>, because a rule has only one move destination.</param>
    /// <param name="stopProcessingRules">Stop evaluating later rules once this one matches.</param>
    /// <param name="enabled">Whether the rule acts on mail immediately. True by default.</param>
    /// <param name="storeId">The mailbox to write to. Defaults to the profile's default delivery store.</param>
    [ServiceAction("create", Destructive = true)]
    MailRuleMutationResult Create(
        string name,
        string? fromAddress = null,
        string? subjectContains = null,
        string? moveToFolder = null,
        string? assignCategories = null,
        bool deleteMessage = false,
        bool stopProcessingRules = false,
        bool enabled = true,
        string? storeId = null);

    /// <summary>
    /// Changes an existing rule's clauses and commits it.
    ///
    /// <para>
    /// A clause that is not given is left exactly as it was. To remove a condition, pass an empty
    /// string for it; to switch an action off, pass false. A rule cannot be updated into having no
    /// conditions or no actions, for the same reason it cannot be created that way.
    /// </para>
    /// </summary>
    /// <param name="name">The rule to change. Must match exactly one rule in the store.</param>
    /// <param name="fromAddress">Replace the sender-address condition. Empty string removes it.</param>
    /// <param name="subjectContains">Replace the subject condition. Empty string removes it.</param>
    /// <param name="moveToFolder">Replace the move destination. Empty string removes the move action.</param>
    /// <param name="assignCategories">Replace the assigned categories, comma-separated. Empty string removes the action.</param>
    /// <param name="deleteMessage">Switch the delete action on or off. Omit to leave it as it is. Because Outlook stores a delete as a move to Deleted Items, switching it off does not necessarily remove that move; clear it by passing an empty string for <paramref name="moveToFolder"/>.</param>
    /// <param name="stopProcessingRules">Switch stop-processing on or off. Omit to leave it as it is.</param>
    /// <param name="newName">Rename the rule. Must not collide with another rule in the store.</param>
    /// <param name="storeId">The mailbox to write to. Defaults to the profile's default delivery store.</param>
    [ServiceAction("update", Destructive = true)]
    MailRuleMutationResult Update(
        string name,
        string? fromAddress = null,
        string? subjectContains = null,
        string? moveToFolder = null,
        string? assignCategories = null,
        bool? deleteMessage = null,
        bool? stopProcessingRules = null,
        string? newName = null,
        string? storeId = null);

    /// <summary>
    /// Switches a rule on or off and commits it.
    ///
    /// <para>
    /// Outlook is explicit that a rule is only enabled once saved, so this is not merely setting a
    /// flag: the store's whole rule collection is rewritten, and until that succeeds the rule's
    /// state has not changed at all.
    /// </para>
    /// </summary>
    /// <param name="name">The rule to switch. Must match exactly one rule in the store.</param>
    /// <param name="enabled">True to have the rule act on mail, false to leave it defined but inert.</param>
    /// <param name="storeId">The mailbox to write to. Defaults to the profile's default delivery store.</param>
    [ServiceAction("set-enabled", Destructive = true)]
    MailRuleMutationResult SetEnabled(string name, bool enabled = true, string? storeId = null);

    /// <summary>
    /// Removes a rule and commits the removal.
    ///
    /// <para>
    /// There is no undo. Prefer <c>set-enabled</c> with false when the intent is to stop a rule
    /// acting rather than to forget what it was.
    /// </para>
    /// </summary>
    /// <param name="name">The rule to remove. Must match exactly one rule in the store.</param>
    /// <param name="storeId">The mailbox to write to. Defaults to the profile's default delivery store.</param>
    [ServiceAction("delete", Destructive = true)]
    MailRuleMutationResult Delete(string name, string? storeId = null);
}
