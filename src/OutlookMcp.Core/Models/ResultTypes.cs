using System.Text.Json.Serialization;

namespace OutlookMcp.Core.Models;

/// <summary>
/// Base result type for all Core operations.
/// Exceptions propagate naturally â€” batch.Execute() re-throws them via TaskCompletionSource.
/// </summary>
public abstract class ResultBase
{
    public bool Success { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FilePath { get; set; }
}

/// <summary>
/// Result for operations that don't return data (create, delete, etc.)
/// </summary>
public class OperationResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Action { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }
}


// â”€â”€ File / Session â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


// â”€â”€ Slide â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€




// â”€â”€ Shape â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€




// â”€â”€ Text â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€




// â”€â”€ Table (in shapes) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


// â”€â”€ Master / Layout â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€




// â”€â”€ Notes â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


// â”€â”€ Transition â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


// â”€â”€ Animation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€



// â”€â”€ Export â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


// â”€â”€ Outlook application / folder / mail â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

public class OutlookApplicationStatusResult : ResultBase
{
    public bool Connected { get; set; }
    public string Version { get; set; } = string.Empty;
    public int ExplorerCount { get; set; }
    public int InspectorCount { get; set; }
    public int StoreCount { get; set; }

    /// <summary>
    /// The classic-vs-new Outlook flavour detected on this machine. Only "classic-desktop" is
    /// supported by this server, since new Outlook for Windows has no COM object model. See #35.
    /// </summary>
    public string OutlookFlavor { get; set; } = string.Empty;

    /// <summary>True if this process is running elevated (as Administrator).</summary>
    public bool ProcessElevated { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CurrentFolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CurrentFolderPath { get; set; }

    public bool HasActiveMailSelection { get; set; }
}

public class OutlookFolderListResult : ResultBase
{
    public List<OutlookFolderInfo> Folders { get; set; } = [];
}

public class OutlookFolderResolveResult : ResultBase
{
    public bool Resolved { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RequestedFolder { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DefaultRole { get; set; }

    /// <summary>
    /// Set when the operation succeeded but something about the result needs saying - most often that
    /// Outlook reports no usable path for the folder, so it cannot be addressed by path afterwards.
    /// A caller that ignores this would silently store an identifier that resolves to nothing.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }

    public int ChildFolderCount { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ItemCount { get; set; }
}

public class OutlookFolderInfo
{
    public string Role { get; set; } = string.Empty;
    public bool Available { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    /// <summary>
    /// The store this folder was read from. Every store in a profile has its own Inbox, so a folder
    /// listing that does not name its store is ambiguous on any profile with more than one - and the
    /// caller has no way to notice. See #38.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreName { get; set; }

    /// <summary>
    /// Why a role is not available, when that needs explaining. Set when a store answers for a
    /// default role it does not actually have - Outlook returns a folder object that is not in the
    /// store's tree and so cannot be addressed. See #38.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ItemCount { get; set; }
}

/// <summary>
/// The stores in the current Outlook profile - the Exchange mailbox, any archives, any PST data
/// files, and any additional accounts. See #38.
/// </summary>
public class OutlookStoreListResult : ResultBase
{
    public List<OutlookStoreInfo> Stores { get; set; } = [];
}

/// <summary>
/// The reminders Outlook is holding, across appointments, tasks and flagged mail.
///
/// <para>
/// The counts matter as much as the rows. Most reminders on a long-lived mailbox are overdue, so a
/// page of rows without the totals reads as the whole picture when it is a small slice of one.
/// </para>
/// </summary>
public class MailReminderListResult : ResultBase
{
    public List<MailReminderInfo> Reminders { get; set; } = [];

    /// <summary>Every reminder Outlook holds, regardless of filtering or <c>maxCount</c>.</summary>
    public int TotalCount { get; set; }

    /// <summary>Reminders due now or later.</summary>
    public int UpcomingCount { get; set; }

    /// <summary>
    /// Reminders whose time has already passed. Usually the large majority, and excluded by default.
    /// </summary>
    public int OverdueCount { get; set; }
}

public class MailReminderInfo
{
    /// <summary>The text Outlook shows in the reminder window - normally the item's subject.</summary>
    public string Caption { get; set; } = string.Empty;

    /// <summary>
    /// When the reminder is set for, taken from <c>OriginalReminderDate</c>. Deliberately not
    /// <c>NextReminderDate</c>, which Outlook leaves at the OLE zero date unless the reminder has
    /// been snoozed.
    /// </summary>
    public DateTime ReminderTime { get; set; }

    /// <summary>
    /// The time a snoozed reminder will fire again, or null when it has not been snoozed.
    /// </summary>
    public DateTime? NextReminderTime { get; set; }

    /// <summary>Whether <see cref="ReminderTime"/> has already passed.</summary>
    public bool IsOverdue { get; set; }

    /// <summary>The kind of item being reminded about: appointment, task, mail, or contact.</summary>
    public string ItemType { get; set; } = string.Empty;

    /// <summary>The subject of the underlying item, when it has one.</summary>
    public string? Subject { get; set; }
}

/// <summary>
/// The mailbox's inbox rules.
///
/// <para>
/// Rules move, delete and forward mail before anything else in this surface sees it. Without them a
/// question like "why is nothing arriving from this sender?" gets a confident empty folder instead of
/// the answer, which is that a rule filed it elsewhere.
/// </para>
/// </summary>
public class MailRuleListResult : ResultBase
{
    public List<MailRuleInfo> Rules { get; set; } = [];
}

public class MailRuleInfo
{
    /// <summary>The name shown in Outlook's rule list.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// False for a rule that exists but is switched off. A disabled rule explains nothing about
    /// where mail went, so the two must not be conflated.
    /// </summary>
    public bool Enabled { get; set; }

    /// <summary>
    /// The order Outlook evaluates rules in. It matters because a rule can stop processing
    /// altogether, so a later rule may never run however well it matches.
    /// </summary>
    public int ExecutionOrder { get; set; }

    /// <summary>
    /// <c>receive</c> for a rule that runs on arriving mail, <c>send</c> for one that runs on
    /// outgoing mail. Never the raw <c>OlRuleType</c> ordinal.
    /// </summary>
    public string RuleType { get; set; } = "receive";

    /// <summary>
    /// True for a rule that only runs in this Outlook client, as opposed to one the server applies.
    /// A client-only rule does nothing while Outlook is closed, which is a common reason mail is
    /// filed late or not at all.
    /// </summary>
    public bool IsLocalRule { get; set; }

    /// <summary>
    /// The conditions actually in use, by name - <c>from</c>, <c>subject</c>, and so on.
    ///
    /// <para>
    /// Only populated when <c>includeDetail</c> is set. Outlook's <c>Conditions</c> collection has a
    /// fixed length covering every condition it supports, so this is the subset with
    /// <c>Enabled</c> set; the raw count would report a one-line rule as having 31 conditions.
    /// </para>
    /// </summary>
    public List<string> Conditions { get; set; } = [];

    /// <summary>
    /// The actions actually in use, by name. Same fixed-collection caveat as <c>Conditions</c>.
    /// </summary>
    public List<string> Actions { get; set; } = [];

    /// <summary>
    /// Where a move-to-folder rule files its mail. Null when the rule does not move anything.
    /// This is the field that turns "your mail is being moved" into something the caller can act on.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MoveToFolderPath { get; set; }

    /// <summary>
    /// The senders a from-rule matches.
    ///
    /// <para>
    /// Rule recipients are stored unresolved, so Outlook leaves <c>Recipient.Address</c> blank and
    /// puts the address in <c>Name</c>. Reading only <c>Address</c> reports every from-rule as
    /// matching nobody.
    /// </para>
    /// </summary>
    public List<string> FromAddresses { get; set; } = [];
}

/// <summary>
/// The mailbox's master category list. Outlook does not validate the string
/// <c>mail set-categories</c> writes, so a category that is not in this list is accepted, reported
/// as a success, and then cannot be filtered or coloured by. This is how a caller finds out which
/// names are real before writing one.
/// </summary>
public class MailCategoryListResult : ResultBase
{
    public List<MailCategoryInfo> Categories { get; set; } = [];
}

public class MailCategoryInfo
{
    /// <summary>
    /// The value to pass to <c>set-categories</c>. Outlook keeps these unique within the list, so
    /// the name - not the id - is how a category is addressed.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// The colour Outlook shows for this category, as a name such as <c>yellow</c> or
    /// <c>darkTeal</c>, never as the raw <c>OlCategoryColor</c> ordinal. An ordinal is not something
    /// a caller can show a user or reason about, and <c>none</c> is a real value meaning the
    /// category was created without a colour.
    /// </summary>
    public string Color { get; set; } = "none";

    /// <summary>
    /// Outlook's stable identifier for the category. Present for completeness; nothing in this
    /// surface takes it, because <c>set-categories</c> works in names.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CategoryId { get; set; }

    /// <summary>
    /// The keyboard shortcut assigned in Outlook, when one is. Null for the common case of no
    /// shortcut rather than a zero that would read as a real key.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShortcutKey { get; set; }
}

public class OutlookStoreInfo
{
    /// <summary>
    /// The id to pass as <c>storeId</c>. Display names are not unique - two accounts can both be
    /// called "Archive" - so this, not the name, is what addresses a store.
    /// </summary>
    public string StoreId { get; set; } = string.Empty;

    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// True for the default delivery store: the mailbox an unqualified request lands in.
    /// </summary>
    public bool IsDefaultStore { get; set; }

    /// <summary>
    /// True for a local data file (PST/OST opened as a data file) rather than a server mailbox.
    /// </summary>
    public bool IsDataFileStore { get; set; }

    /// <summary>
    /// Outlook's own classification: <c>primaryExchangeMailbox</c>, <c>deltaSyncMailbox</c>,
    /// <c>publicFolders</c>, <c>notExchange</c>, and so on. Null when the store declines to say.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ExchangeStoreType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FilePath { get; set; }

    /// <summary>
    /// The address of the account that delivers to this store, when one does. Null for a store no
    /// account delivers to - an archive or an imported data file.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AccountSmtpAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AccountDisplayName { get; set; }

    /// <summary>
    /// The store's root folder path, which is what a <c>folder</c> argument must start with to
    /// address anything inside this store by path.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RootFolderPath { get; set; }
}

public class OutlookFolderItemListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    public int TotalItemCount { get; set; }
    public int ReturnedCount { get; set; }

    /// <summary>
    /// True when <c>maxCount</c> stopped this listing short of the folder. A caller MUST NOT read
    /// the returned items as the folder's full contents when this is set. See #91.
    /// </summary>
    public bool Truncated { get; set; }

    /// <summary>
    /// The property items were ordered by before the cap was applied - <c>receivedTime</c> for mail
    /// folders, <c>lastModificationTime</c> for folders whose items have no received time
    /// (calendars, contacts). Null only when the store refused to sort at all, in which case the
    /// order is arbitrary and a truncated listing is an arbitrary subset - which is why it is
    /// reported rather than left for a caller to assume. See #91.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SortedBy { get; set; }

    /// <summary>Direction of <see cref="SortedBy"/>: <c>descending</c> - newest first.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SortDirection { get; set; }

    public List<OutlookFolderItemInfo> Items { get; set; } = [];
}

public class OutlookFolderItemInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MessageClass { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Preview { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? ReceivedTime { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Start { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? End { get; set; }

    public bool Unread { get; set; }
}

public class ActiveMailResult : ResultBase
{
    public bool HasActiveMail { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? To { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Cc { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Bcc { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SenderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SenderEmailAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CurrentFolderPath { get; set; }

    /// <summary>
    /// Identifier of the thread this message belongs to. Pass it - or simply this message's entry id -
    /// to <c>mail.get-conversation</c> to retrieve the whole thread. Null on stores that do not
    /// support conversation view. See #39.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ConversationId { get; set; }

    /// <summary>
    /// The thread's topic: the original subject with reply and forward prefixes stripped, which is
    /// why it is reported separately from <see cref="Subject"/>. See #39.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ConversationTopic { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BodyPreview { get; set; }

    public List<string> Categories { get; set; } = [];

    /// <summary>
    /// Follow-up flag state: <c>none</c>, <c>flagged</c> or <c>complete</c>. Always populated so
    /// "not flagged" is distinguishable from "flags not reported here". See #15.
    /// </summary>
    public string FlagStatus { get; set; } = "none";

    /// <summary>The flag's label, e.g. "Follow up". Absent when nothing is flagged.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FlagRequest { get; set; }

    /// <summary>When the follow-up is due, or absent when it has no date.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? FlagDueDate { get; set; }

    /// <summary>
    /// Names of properties whose read was blocked by the Outlook Object Model Guard rather than
    /// the property simply being absent. See <see cref="MailSummaryInfo.AccessDenied"/>. #30.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? AccessDenied { get; set; }

    public bool Unread { get; set; }
    public int Importance { get; set; }
    public int AttachmentCount { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? ReceivedTime { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? SentOn { get; set; }
}

public class MailDraftResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? To { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Cc { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Bcc { get; set; }

    public bool Displayed { get; set; }
    public bool Saved { get; set; }
    public int BodyLength { get; set; }
}

public class MailListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Query { get; set; }

    public int TotalItemCount { get; set; }
    public int ReturnedCount { get; set; }

    /// <summary>
    /// Number of folder items actually scanned/matched against by this call (i.e. items
    /// Outlook's index/filter evaluated, not the client-side substring fallback's old fixed cap).
    /// </summary>
    public int ScannedCount { get; set; }

    /// <summary>
    /// Items that were scanned but could not be summarised at all - a folder item of a type this
    /// surface does not model, for instance.
    ///
    /// <para>
    /// Reported rather than left implicit because a listing whose numbers do not add up is how
    /// "here is what is in your folder" quietly becomes false. Meeting requests used to land here by
    /// accident - invisibly, since nothing was counted either - which is the bug #32 records.
    /// </para>
    /// </summary>
    public int SkippedItemCount { get; set; }

    /// <summary>
    /// True when this call did not exhaustively scan/match every item in
    /// <see cref="TotalItemCount"/> -- either because the result-count cap (<c>maxCount</c>) was
    /// reached, or (for <c>mail.search</c>'s client-side body-substring fallback path only) a
    /// bounded scan limit was hit before exhausting the folder. A client MUST NOT read an empty
    /// or short <see cref="Messages"/> list as "no such mail exists" when this is true -- there
    /// may be more matches beyond what was scanned/returned. See #27.
    /// </summary>
    public bool Truncated { get; set; }

    /// <summary>
    /// Opaque continuation token for the next page, or <see langword="null"/> when this response
    /// reached the end of the result set. Pass it back unchanged as <c>cursor</c> on an otherwise
    /// identical call. See <see cref="HasMore"/> for the condition to loop on, and #43.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? NextCursor { get; set; }

    /// <summary>
    /// True when a further page can be retrieved with <see cref="NextCursor"/>. This is the flag to
    /// drive a paging loop with; <see cref="Truncated"/> only reports that this call stopped early
    /// and says nothing about whether continuing is possible.
    /// </summary>
    public bool HasMore { get; set; }

    /// <summary>
    /// The property results are ordered by. Paging is a keyset walk over this ordering rather than a
    /// numeric offset, so it is stated explicitly instead of left as an implementation detail a
    /// caller has to infer (#43).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SortedBy { get; set; }

    /// <summary>Direction of <see cref="SortedBy"/>: <c>descending</c> (newest first).</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SortDirection { get; set; }

    /// <summary>
    /// Which engine answered this query: <c>clientScan</c> (each candidate hydrated and checked
    /// client-side, substring semantics, bounded by a scan limit) or <c>contentIndex</c> (Outlook's
    /// full-text index, whole-word semantics, no scan horizon).
    ///
    /// <para>
    /// Reported on every search because an empty result means different things depending on which
    /// one produced it, and a caller cannot tell them apart otherwise. See #42.
    /// </para>
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SearchEngine { get; set; }

    /// <summary>
    /// Anything the caller needs to know about how the answer was arrived at that the numbers do not
    /// say - most importantly, that a requested content-index search could not be served by this
    /// store and fell back to the client-side scan.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    public List<MailSummaryInfo> Messages { get; set; } = [];
}

/// <summary>
/// One mail thread: every message in the conversation, across folders, in reading order (#39).
/// </summary>
public class MailConversationResult : ResultBase
{
    /// <summary>Identifier of the thread, matching <c>conversationId</c> on read and list results.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ConversationId { get; set; }

    /// <summary>The thread's topic: the original subject with reply/forward prefixes stripped.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ConversationTopic { get; set; }

    /// <summary>
    /// False when the store cannot provide a conversation view at all (some PST and third-party
    /// stores). Reported explicitly, alongside <c>success: false</c>, rather than returned as an
    /// empty-but-successful thread: "this message has no replies" and "this store cannot tell you
    /// whether it has replies" are different answers and must not look alike.
    /// </summary>
    public bool ConversationSupported { get; set; } = true;

    /// <summary>Number of messages in the whole thread, before <c>maxCount</c> is applied.</summary>
    public int TotalItemCount { get; set; }

    /// <summary>Number of messages actually returned in <see cref="Messages"/>.</summary>
    public int ReturnedCount { get; set; }

    /// <summary>
    /// Number of thread entries that were not mail items - a meeting request or a delivery report
    /// filed into the same conversation - and so were counted but not returned. Reported rather than
    /// silently dropped, so a caller can see why the counts differ.
    /// </summary>
    public int SkippedItemCount { get; set; }

    /// <summary>
    /// True when <c>maxCount</c> cut the thread short. A caller MUST NOT read a truncated thread as
    /// the whole conversation.
    /// </summary>
    public bool Truncated { get; set; }

    /// <summary>The property items are ordered by: <c>receivedTime</c>.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SortedBy { get; set; }

    /// <summary>Direction of <see cref="SortedBy"/>: <c>ascending</c> - oldest first, reading order.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SortDirection { get; set; }

    public List<MailSummaryInfo> Messages { get; set; } = [];

    /// <summary>
    /// Members of the thread that are not mail: meeting invitations, the calendar appointments they
    /// create, and acceptances or declines. These used to be reduced to a number, which on a real
    /// thread hid most of it - a seven-item conversation reported three messages and the digit 4.
    /// The invitation and the acceptance are often the substance of a thread, so they are named
    /// here rather than counted. See #111.
    /// </summary>
    public List<MailThreadItemInfo> OtherItems { get; set; } = [];
}

/// <summary>
/// A non-mail member of a conversation. Deliberately thinner than <see cref="MailSummaryInfo"/>:
/// these items have no sender or recipients in the mail sense, and presenting empty fields for them
/// would suggest the data was missing rather than inapplicable.
/// </summary>
public class MailThreadItemInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    /// <summary>
    /// What this is, as a name a caller can act on: <c>appointment</c>, <c>meetingRequest</c>,
    /// <c>meetingResponse</c>, <c>task</c>, <c>contact</c>, or <c>unknown</c>. Never a raw class
    /// ordinal - a number here would be exactly the opacity this field exists to remove.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    /// <summary>Which folder it lives in - a thread spans folders, including Calendar.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    /// <summary>
    /// When it happened, used for the same ordering as the messages. For an appointment this is its
    /// start; for a meeting request or response it is when it was sent or received.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Timestamp { get; set; }
}

public class MailSummaryInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SenderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SenderEmailAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? To { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Cc { get; set; }

    /// <summary>
    /// Identifier of the thread this message belongs to, so a caller can group a listing into
    /// threads, or fetch one, without a separate read per message. See #39.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ConversationId { get; set; }

    /// <summary>The thread's topic: the subject with reply/forward prefixes stripped. See #39.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ConversationTopic { get; set; }

    /// <summary>
    /// Folder this message lives in. Populated for thread results, where items genuinely span
    /// folders (a reply sits in Sent Items while the original sits in the Inbox), and omitted for a
    /// folder listing, where the folder is already on the envelope. See #39.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BodyPreview { get; set; }

    public List<string> Categories { get; set; } = [];

    /// <summary>
    /// Names of properties whose read was blocked by the Outlook Object Model Guard (a security
    /// prompt was shown and not approved, or Outlook aborted the call outright) rather than the
    /// property simply being absent. A client seeing e.g. <c>senderEmailAddress: null</c> plus
    /// <c>"senderEmailAddress"</c> in this list should not treat that as "no sender" -- it means
    /// access was denied. Empty when no property read was blocked. See #30 (Rule 22: security
    /// denials must never be silently indistinguishable from "value not present").
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? AccessDenied { get; set; }

    /// <summary>
    /// What kind of item this is: <c>mail</c>, <c>meetingRequest</c>, <c>meetingCancellation</c>,
    /// <c>meetingResponse</c> or <c>other</c>.
    ///
    /// <para>
    /// A meeting invitation is a <c>MeetingItem</c>, not a <c>MailItem</c>, and listings used to drop
    /// it silently. It is now listed - but a caller must be able to tell it apart, because the two
    /// afford completely different actions: replying to an invitation is not accepting it. See #32.
    /// </para>
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    /// <summary>
    /// Follow-up flag state: <c>none</c>, <c>flagged</c> or <c>complete</c>. Always populated, never
    /// omitted, so a caller can tell "this message is not flagged" from "this listing does not report
    /// flags" - the two mean very different things when deciding what still needs attention. See #15.
    /// </summary>
    public string FlagStatus { get; set; } = "none";

    /// <summary>The flag's label, e.g. "Follow up" or "Review". Absent when nothing is flagged.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FlagRequest { get; set; }

    /// <summary>
    /// When the follow-up is due. Absent when the message is unflagged or was flagged without a date;
    /// Outlook's far-future sentinel for "no date" is reported as absent rather than as a real date.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? FlagDueDate { get; set; }

    public bool Unread { get; set; }
    public bool IsDraft { get; set; }
    public int Importance { get; set; }
    public int AttachmentCount { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? ReceivedTime { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? SentOn { get; set; }
}

public class MailSendResult : ResultBase
{
    public bool Sent { get; set; }

    /// <summary>
    /// True when the outcome of this send request could not be determined (e.g. the underlying
    /// operation timed out while a security prompt was on screen). An indeterminate outcome is
    /// deliberately NOT the same as <c>Success = false</c>: the mail may have actually sent. A
    /// client seeing <c>indeterminate: true</c> must not blindly retry -- retrying an
    /// already-sent message would duplicate it. Re-check via <c>mail.read</c> using
    /// <see cref="EntryId"/>/<see cref="StoreId"/> (if known) before deciding whether to resend.
    /// See #29.
    /// </summary>
    public bool Indeterminate { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? To { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? SentOn { get; set; }
}

public class MailMutationResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? To { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Cc { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Bcc { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    public List<string> Categories { get; set; } = [];

    public bool Deleted { get; set; }
    public bool Moved { get; set; }
    public bool Read { get; set; }
}

public class CalendarListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Start { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? End { get; set; }

    public int TotalItemCount { get; set; }
    public int ReturnedCount { get; set; }

    /// <summary>
    /// Whether occurrences of recurring series were expanded into the list. Expansion needs a bounded
    /// range, since a series with no end date has infinitely many occurrences. When this is false the
    /// list contains series masters only, and a recurring meeting will be missing from every date but
    /// its first.
    /// </summary>
    public bool RecurringExpanded { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    public List<CalendarSummaryInfo> Appointments { get; set; } = [];
}

public class CalendarSummaryInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Location { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Organizer { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BodyPreview { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Start { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? End { get; set; }

    public bool AllDay { get; set; }
    public bool ReminderSet { get; set; }
    public int BusyStatus { get; set; }

    public bool IsRecurring { get; set; }

    /// <summary>
    /// <c>notRecurring</c>, <c>master</c>, <c>occurrence</c> or <c>exception</c>. An
    /// <c>occurrence</c> is one instance of a series and carries the master's entry id, so editing it
    /// by entry id edits the whole series.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RecurrenceState { get; set; }
}

/// <summary>
/// A recurrence pattern, as Outlook stores it.
/// </summary>
public class RecurrencePatternInfo
{
    /// <summary>
    /// <c>daily</c>, <c>weekly</c>, <c>monthly</c>, <c>monthNth</c>, <c>yearly</c> or
    /// <c>yearNth</c>.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RecurrenceType { get; set; }

    /// <summary>How many units between occurrences - every 2 weeks, every 3 days.</summary>
    public int Interval { get; set; }

    /// <summary>Lower-case day names, for weekly patterns.</summary>
    public List<string> DaysOfWeek { get; set; } = [];

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? DayOfMonth { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? MonthOfYear { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? PatternStartDate { get; set; }

    /// <summary>Meaningless when <c>NoEndDate</c> is true - Outlook still stores a sentinel there.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? PatternEndDate { get; set; }

    public bool NoEndDate { get; set; }

    /// <summary>How many occurrences the series has, when it is bounded by a count.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Occurrences { get; set; }

    /// <summary>Length of a single occurrence, in minutes.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? DurationMinutes { get; set; }

    /// <summary>
    /// How many occurrences differ from the pattern - moved, shortened or cancelled. Non-zero means
    /// the pattern alone does not describe the series.
    /// </summary>
    public int ExceptionCount { get; set; }
}

public class CalendarItemResult : ResultBase
{
    public bool HasItem { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Location { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Organizer { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BodyPreview { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Start { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? End { get; set; }

    public bool AllDay { get; set; }
    public bool ReminderSet { get; set; }
    public int BusyStatus { get; set; }

    /// <summary>
    /// True when the item is a meeting - it has attendees and an organiser - rather than a private
    /// appointment. The two afford different actions, so a caller must be able to tell them apart.
    /// </summary>
    public bool IsMeeting { get; set; }

    /// <summary>
    /// Everybody invited, with the response each has given so far. Empty for a plain appointment.
    /// </summary>
    public List<MeetingAttendeeInfo> Attendees { get; set; } = [];

    public bool IsRecurring { get; set; }

    /// <summary>
    /// <c>notRecurring</c>, <c>master</c>, <c>occurrence</c> or <c>exception</c>.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RecurrenceState { get; set; }

    /// <summary>
    /// The series pattern. Null for a non-recurring item - absence means "not a series", not
    /// "a series whose pattern could not be read", which would be reported as a failure instead.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public RecurrencePatternInfo? Recurrence { get; set; }
}

/// <summary>
/// One invitee on a meeting.
/// </summary>
public class MeetingAttendeeInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Address { get; set; }

    /// <summary>
    /// <c>required</c>, <c>optional</c>, <c>resource</c> or <c>organizer</c>.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Type { get; set; }

    /// <summary>
    /// <c>none</c>, <c>organizer</c>, <c>tentative</c>, <c>accepted</c>, <c>declined</c> or
    /// <c>notResponded</c>. <c>none</c> means the item is not a meeting, not that they declined.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ResponseStatus { get; set; }

    /// <summary>
    /// Whether Outlook could resolve the name against an address book or as a valid SMTP address.
    /// An unresolved attendee will never receive the invitation.
    /// </summary>
    public bool Resolved { get; set; }
}

public class CalendarAppointmentResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Location { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Start { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? End { get; set; }

    public bool Saved { get; set; }
    public bool Displayed { get; set; }
    public bool AllDay { get; set; }

    /// <summary>
    /// True when attendees were named, so Outlook stored a meeting rather than a private appointment.
    /// </summary>
    public bool IsMeeting { get; set; }

    /// <summary>
    /// Whether an invitation was actually sent. Creating a meeting saves it to the caller's own
    /// calendar and tells nobody; only <c>sendInvitation</c> mails the attendees.
    /// </summary>
    public bool InvitationSent { get; set; }

    /// <summary>
    /// Attendees as Outlook resolved them.
    /// </summary>
    public List<MeetingAttendeeInfo> Attendees { get; set; } = [];

    /// <summary>
    /// Attendees Outlook could not resolve. Non-empty means the meeting was not created: an
    /// unresolvable attendee never receives the invitation, so saving anyway would report success
    /// for a meeting that cannot reach the person the caller named.
    /// </summary>
    public List<string> UnresolvedAttendees { get; set; } = [];

    /// <summary>
    /// True when a recurrence pattern was applied, so the item is a series master rather than a
    /// single appointment.
    /// </summary>
    public bool IsRecurring { get; set; }

    /// <summary>
    /// The pattern as Outlook stored it, read back after saving rather than echoed from the request.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public RecurrencePatternInfo? Recurrence { get; set; }
}

/// <summary>
/// One person's availability over the requested window.
/// </summary>
public class FreeBusyPersonInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Address { get; set; }

    /// <summary>
    /// Whether Outlook resolved the name. An unresolved person's availability is unknown, never free.
    /// </summary>
    public bool Resolved { get; set; }

    /// <summary>
    /// Outlook's raw slot string: one character per interval, <c>0</c> free, <c>1</c> tentative,
    /// <c>2</c> busy, <c>3</c> out of office, <c>4</c> working elsewhere.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Availability { get; set; }

    /// <summary>
    /// The same information as merged non-free intervals, which is what a caller looking for a slot
    /// actually needs. Free time is everything these do not cover.
    /// </summary>
    public List<FreeBusyPeriodInfo> BusyPeriods { get; set; } = [];
}

/// <summary>
/// A stretch of non-free time.
/// </summary>
public class FreeBusyPeriodInfo
{
    public DateTimeOffset Start { get; set; }
    public DateTimeOffset End { get; set; }

    /// <summary>
    /// <c>tentative</c>, <c>busy</c>, <c>outOfOffice</c>, <c>workingElsewhere</c> or <c>unknown</c>.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Status { get; set; }
}

public class CalendarFreeBusyResult : ResultBase
{
    public DateTimeOffset Start { get; set; }
    public DateTimeOffset End { get; set; }

    /// <summary>
    /// Minutes covered by each character of <see cref="FreeBusyPersonInfo.Availability"/>.
    /// </summary>
    public int IntervalMinutes { get; set; }

    public List<FreeBusyPersonInfo> People { get; set; } = [];

    /// <summary>
    /// Set when the answer covers less time than was asked for.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    /// <summary>
    /// Attendees Outlook could not resolve. Non-empty means the lookup failed: reporting an
    /// unresolvable person as free would schedule over a calendar nobody ever looked at.
    /// </summary>
    public List<string> UnresolvedAttendees { get; set; } = [];
}

/// <summary>
/// The outcome of answering a meeting invitation.
/// </summary>
public class MeetingResponseResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    /// <summary>
    /// <c>accept</c>, <c>decline</c> or <c>tentative</c>, as it was applied.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Response { get; set; }

    /// <summary>
    /// Whether the organiser was told. Answering updates your own calendar either way; only
    /// <c>sendResponse</c> mails them.
    /// </summary>
    public bool ResponseSent { get; set; }
}

public class CalendarMutationResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }


    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Location { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? Start { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DateTimeOffset? End { get; set; }

    public bool AllDay { get; set; }
    public bool Updated { get; set; }
    public bool Deleted { get; set; }

    /// <summary>
    /// What the change actually touched: <c>series</c> for the whole item (recurring or not) or
    /// <c>occurrence</c> for a single instance of a series.
    ///
    /// <para>
    /// This is reported rather than inferred because an occurrence carries its master's entry id, so
    /// a series-wide edit and a single-instance edit are indistinguishable in the response otherwise.
    /// A caller that cancelled one stand-up needs to be able to see that it cancelled one stand-up.
    /// </para>
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Scope { get; set; }
}

public class AttachmentListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    public int AttachmentCount { get; set; }
    public List<AttachmentInfo> Attachments { get; set; } = [];
}

public class AttachmentInfo
{
    public int Index { get; set; }
    public string FileName { get; set; } = string.Empty;
    public int SizeBytes { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContentId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DisplayName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Type { get; set; }

    public bool Hidden { get; set; }
}

public class AttachmentSaveResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    public int SavedCount { get; set; }
    public List<string> SavedFiles { get; set; } = [];

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }
}

public class AttachmentMutationResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    public int AttachmentCount { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FileName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }
}

// -- Contact ------------------------------------------------------------

/// <summary>
/// A Contacts folder listing.
///
/// <para>
/// <see cref="Contacts"/>, <see cref="DistributionLists"/> and <see cref="SkippedItemCount"/>
/// together always account for <see cref="ScannedItemCount"/>. That identity is the contract: a
/// Contacts folder holds distribution lists as well as people, and an earlier implementation
/// dropped anything that was not a <c>ContactItem</c> while still reporting the full folder size,
/// so a caller was told 83 and handed 82 with nothing to indicate the difference.
/// </para>
/// </summary>
public class ContactListResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    /// <summary>Items in the folder, whether or not this call looked at them.</summary>
    public int TotalItemCount { get; set; }

    /// <summary>Items this call actually examined. Below <see cref="TotalItemCount"/> when truncated.</summary>
    public int ScannedItemCount { get; set; }

    /// <summary>Contacts returned, equal to <see cref="Contacts"/> count.</summary>
    public int ReturnedCount { get; set; }

    /// <summary>Scanned items that could not be read at all. Not a count of non-contact items.</summary>
    public int SkippedItemCount { get; set; }

    public bool Truncated { get; set; }

    public List<ContactSummaryInfo> Contacts { get; set; } = [];

    /// <summary>Distribution lists in the same folder. Groups of people, not people.</summary>
    public List<ContactDistributionListInfo> DistributionLists { get; set; } = [];
}

public class ContactSummaryInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    /// <summary>
    /// A label that is never blank: full name, else company, else an email address, else a
    /// placeholder naming the fact. Some real contacts carry no name at all, and a row with nothing
    /// to show is a row the caller cannot act on or display.
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>The raw Outlook <c>FullName</c>, which is genuinely empty for some contacts.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FullName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LastName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CompanyName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Email1Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BusinessTelephoneNumber { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MobileTelephoneNumber { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BodyPreview { get; set; }
}

/// <summary>
/// A distribution list found in a Contacts folder. Returned so that it is visible rather than
/// silently discarded; its members are not expanded.
/// </summary>
public class ContactDistributionListInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    public string DisplayName { get; set; } = string.Empty;

    public int MemberCount { get; set; }
}

public class ContactItemResult : ResultBase
{
    /// <summary>
    /// False when no contact could be resolved. This is not an error: asking for the active contact
    /// when none is open is a legitimate question with the answer "none".
    /// </summary>
    public bool HasItem { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    /// <summary>See <see cref="ContactSummaryInfo.DisplayName"/>. Empty only when there is no item.</summary>
    public string DisplayName { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FullName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FirstName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LastName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CompanyName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? JobTitle { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Email1Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Email2Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BusinessTelephoneNumber { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MobileTelephoneNumber { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BodyPreview { get; set; }
}

public class ContactMutationResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DisplayName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FullName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CompanyName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? JobTitle { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Email1Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Email2Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BusinessTelephoneNumber { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MobileTelephoneNumber { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FolderPath { get; set; }

    public bool Saved { get; set; }
    public bool Displayed { get; set; }
    public bool Updated { get; set; }
    public bool Deleted { get; set; }
}

/// <summary>
/// What the user is looking at in the main Outlook window. Reported so an agent can act on the
/// current context instead of guessing a folder name.
/// </summary>
public class OutlookExplorerContextResult : ResultBase
{
    /// <summary>
    /// False when Outlook is running with no explorer window. That is not an error: it is a real
    /// state, reached by closing every window while an add-in keeps the process alive.
    /// </summary>
    public bool HasExplorer { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CurrentFolderName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CurrentFolderPath { get; set; }

    /// <summary>How many items are selected. The remaining fields describe the first of them.</summary>
    public int SelectionCount { get; set; }

    /// <summary>True when the first selected item is a mail item, so mail actions apply to it.</summary>
    public bool HasMailSelection { get; set; }

    /// <summary>
    /// A readable kind for the first selected item - "mail", "appointment", "contact" and so on -
    /// derived from the item's Outlook object class. Deliberately not the runtime type name: the
    /// wrapper reports itself as <c>__ComObject</c>, which tells a caller nothing.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SelectedItemType { get; set; }

    /// <summary>The raw Outlook message class, such as <c>IPM.Note</c>. Null when nothing is selected.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SelectedItemMessageClass { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SelectedItemSubject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SelectedItemEntryId { get; set; }
}

/// <summary>
/// What is open in the foreground item window, if anything.
/// </summary>
public class OutlookInspectorContextResult : ResultBase
{
    /// <summary>
    /// False when no item window is open. Not an error: "what is open?" has the legitimate answer
    /// "nothing", and reporting that as a failure makes an agent think the mailbox is broken.
    /// </summary>
    public bool HasInspector { get; set; }

    /// <summary>See <see cref="OutlookExplorerContextResult.SelectedItemType"/>.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ItemType { get; set; }

    /// <summary>The raw Outlook message class, such as <c>IPM.Note</c>.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MessageClass { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    /// <summary>
    /// Null until the open item has been saved at least once. An item the user is still composing
    /// has no entry id, so it cannot be addressed by any other action in this surface. Check
    /// <see cref="IsSaved"/> to tell "no id yet" from "id withheld".
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EntryId { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoreId { get; set; }

    /// <summary>
    /// False for an item the user is still composing. Such an item has no entry id, so it cannot
    /// be addressed by any other action in this surface until it is saved.
    /// </summary>
    public bool IsSaved { get; set; }

    /// <summary>
    /// The folder holding the item. On an unsaved compose window Outlook reports the Outbox, which
    /// is where it would go, not where it is - so do not present this as a stored location unless
    /// <see cref="IsSaved"/> is true.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CurrentFolderPath { get; set; }

    /// <summary>The window caption, which is what the user sees in the task bar.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Caption { get; set; }
}