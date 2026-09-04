using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Mail;

[ServiceCategory("mail")]
[McpTool("mail", Title = "Outlook Mail Operations", Destructive = true, Category = "mail",
    Description = "Read the currently active Outlook mail item, list and search mailbox content, and create or send Outlook draft emails without opening a persistent session. "
    + "Use read-active to inspect the currently selected or opened mail item. "
    + "Use read to inspect an explicit Outlook mail item by entry id/store id or fall back to the active mail item. "
    + "Use list and search to inspect the current folder or a default Outlook folder role such as inbox or drafts. "
    + "Both accept structured filters that Outlook evaluates itself before any item is read - unreadOnly, fromAddress, "
    + "subjectContains, receivedAfter, receivedBefore and hasAttachment - which combine with AND. Prefer them over "
    + "listing a folder and filtering client-side, because they reach matches a plain listing would never scan as far as. "
    + "Dates are ISO 8601 ('2024-03-07' or '2024-03-07T14:30'); a bare date means local midnight. "
    + "search's query is a free-text match over subject, sender and the full message body, applied after the structured "
    + "filters. By default it is exhaustive rather than indexed, so it will not miss a term buried deep in a long "
    + "message, but reaching the body means opening every candidate item and the scan stops at a safety limit. "
    + "Set searchMode to 'fullText' to have Outlook's content index answer the query instead: it examines nothing "
    + "client-side, has no scan limit, and so finds matches arbitrarily far back in a large folder - but it matches "
    + "whole words, not substrings, so it finds 'foo' in 'a foo arrived' and not inside 'foobar'. The default is "
    + "'clientScan'. Every search response reports searchEngine ('clientScan' or 'contentIndex'), because an empty "
    + "result means different things depending on which engine produced it; if the index was asked for and the store "
    + "could not serve it, searchEngine says clientScan and message explains why. "
    + "list and search page with an opaque cursor: when a response has hasMore true, pass its nextCursor back as "
    + "cursor on an otherwise identical call to get the next page, and keep going until hasMore is false. Never treat "
    + "a truncated or empty page as proof that no such mail exists while hasMore is true. A cursor only continues the "
    + "exact folder, query and filters that produced it - changing any of them requires restarting without a cursor - "
    + "but maxCount may be changed freely mid-walk. "
    + "Use get-conversation to retrieve a whole mail thread in one call - every message in the "
    + "conversation, oldest first, spanning folders, so a reply in Sent Items comes back alongside "
    + "the original in the Inbox. Prefer it over guessing which search results belong together by "
    + "subject, which is unreliable and cannot see replies filed in other folders. A thread is not "
    + "all mail: meeting invitations, the appointments they create and acceptances come back in "
    + "otherItems with a named type, so treat those as part of the conversation rather than reading "
    + "messages alone as the whole thread. read, list and "
    + "search all report conversationId, so a thread can be reached from any of them. "
    + "Use create-draft to create and save a new Outlook draft with optional recipients, subject, and body text. "
    + "Use reply, reply-all, and forward to create saved draft responses targeting an explicit mail item by entry id/store id, "
    + "or fall back to the active mail item; works headlessly with no Outlook window focused when entryId is supplied. "
    + "forward accepts recipients since a forwarded message otherwise has nobody to send to. All three accept an optional "
    + "body to prepend above the quoted original message. "
    + "Use set-subject, set-body, and set-recipients to edit an existing draft before sending. "
    + "Use set-flag to raise, complete or clear a follow-up flag. flagStatus is 'flagged' (the default), "
    + "'complete' or 'none'; an unrecognised value is refused rather than guessed at. 'complete' and 'none' "
    + "are different outcomes and not interchangeable: 'complete' records that the item was dealt with, "
    + "while 'none' says it was never flagged at all. Optionally pass dueDate and a flagRequest label such "
    + "as 'Review'. Outlook only allows a flag to be completed on a message that has been sent or received, "
    + "so completing a draft is refused with an explanation. read, list and search all report flagStatus, "
    + "always - an unflagged message reports 'none' rather than omitting the field - plus flagRequest and "
    + "flagDueDate when set, so 'what still needs follow-up' costs no extra call per message. "
    + "Use set-categories to apply Outlook categories, and list-categories to discover which ones exist "
    + "before writing any. Outlook does not validate the string set-categories writes: a name that is not "
    + "in the mailbox's list is accepted and reported as a success, then turns out to be uncolourable and "
    + "unfilterable, so discover names rather than guessing them. Colours come back as names such as "
    + "'yellow', never as raw enum numbers. Use create-category to add a category to the master list with "
    + "a colour before assigning it, update-category to recolour, rename or reshortcut one, and "
    + "delete-category to remove one; colours are passed as friendly names too, and an omitted or "
    + "unrecognised colour creates the category with no colour and says so. "
    + "Use list-reminders to see what Outlook is set to remind the user about, earliest first, across "
    + "appointments, tasks and flagged mail. Overdue reminders are excluded by default because on a "
    + "long-lived mailbox they are usually the large majority and bury the ones still to come; the "
    + "result always reports how many were held back. "
    + "Use list-rules to see the mailbox's inbox rules. Rules move, delete and forward mail before this "
    + "tool ever sees it, so when a folder looks empty or a sender's mail is missing, check the rules "
    + "before concluding nothing arrived. Pass includeDetail to get each rule's conditions, actions and "
    + "move-to destination; it is off by default because gathering it is roughly forty times the work. "
    + "list and search also accept flaggedOnly, which Outlook evaluates over the folder rather than the "
    + "server scanning it, so 'show me my outstanding follow-ups' is one cheap call. It returns only "
    + "outstanding flags: a completed flag is finished work and is excluded. "
    + "create-draft, reply, reply-all, forward and set-body all accept bodyFormat, which is 'plain' (the default) "
    + "or 'html'. Pass 'html' when the body argument is markup you want rendered - lists, links, emphasis, tables. "
    + "Leave it as 'plain' for ordinary text: plain text is escaped rather than interpreted, so a body containing "
    + "'<' or '&' arrives exactly as written instead of being silently mangled. An unrecognised value is refused "
    + "rather than guessed at. On reply, reply-all and forward the body goes above the quoted original and the "
    + "quoted message keeps its own formatting either way. "
    + "Use send to send a saved draft by entry id or the current active draft explicitly. Send requires confirm=true "
    + "(it is refused otherwise) and accepts an optional operationId so a retried call with the same operationId after "
    + "a timeout or crash is answered from a cached result instead of risking a duplicate send. "
    + "Set display=true on draft-producing actions to show the draft inspector after saving. "
    + "Use respond-to-meeting to accept, decline or tentatively accept a meeting invitation - a listing's itemType "
    + "says which items are invitations. Responding updates your own calendar; the organiser is told only when "
    + "sendResponse is true, so accepting quietly and notifying them are separate choices. "
    + "Replying to an invitation with reply or forward is not the same thing and does not answer it.")]
public interface IMailCommands
{
    [ServiceAction("read-active", Destructive = false)]
    ActiveMailResult ReadActive();

    [ServiceAction("read", Destructive = false)]
    ActiveMailResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("list", Destructive = false)]
    MailListResult List(
        string? folder = null,
        int maxCount = 25,
        bool unreadOnly = false,
        bool includeBodyPreview = false,
        string? fromAddress = null,
        string? subjectContains = null,
        string? receivedAfter = null,
        string? receivedBefore = null,
        bool? hasAttachment = null,
        bool flaggedOnly = false,
        string? cursor = null);

    [ServiceAction("search", Destructive = false)]
    MailListResult Search(
        string query,
        string? folder = null,
        int maxCount = 25,
        bool unreadOnly = false,
        bool includeBodyPreview = false,
        string? fromAddress = null,
        string? subjectContains = null,
        string? receivedAfter = null,
        string? receivedBefore = null,
        bool? hasAttachment = null,
        bool flaggedOnly = false,
        string? cursor = null,
        string? searchMode = null);

    [ServiceAction("get-conversation", Destructive = false)]
    MailConversationResult GetConversation(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        int maxCount = 50,
        bool includeBodyPreview = false);

    [ServiceAction("respond-to-meeting", Destructive = true)]
    MeetingResponseResult RespondToMeeting(
        string? entryId = null,
        string? storeId = null,
        string response = "accept",
        bool sendResponse = false,
        string? responseText = null,
        bool useActiveMail = false);

    [ServiceAction("create-draft")]
    MailDraftResult CreateMailDraft(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? subject = null,
        string? body = null,
        bool display = false,
        string bodyFormat = "plain");

    [ServiceAction("reply")]
    MailDraftResult Reply(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false,
        string bodyFormat = "plain");

    [ServiceAction("reply-all")]
    MailDraftResult ReplyAll(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false,
        string bodyFormat = "plain");

    [ServiceAction("forward")]
    MailDraftResult Forward(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? body = null,
        bool display = false,
        string bodyFormat = "plain");

    [ServiceAction("send", Destructive = true)]
    MailSendResult Send(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool confirm = false,
        string? operationId = null);

    [ServiceAction("move")]
    MailMutationResult Move(
        string targetFolder,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    /// <summary>
    /// Saves a mail item to disk with <c>MailItem.SaveAs</c>.
    ///
    /// <para>
    /// <paramref name="filePath"/> must be absolute: Outlook accepts a relative path and resolves it
    /// against its own working directory, so the file would land somewhere the caller never looks.
    /// The format defaults to whatever the extension names, and <c>msg</c> always means the Unicode
    /// variant - the ANSI one silently replaces any character outside the machine's code page with
    /// <c>?</c>. An existing file is never replaced unless <paramref name="overwrite"/> is set.
    /// </para>
    /// </summary>
    [ServiceAction("export")]
    ItemExportResult Export(
        string filePath,
        string? format = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool overwrite = false);

    [ServiceAction("delete")]
    MailMutationResult Delete(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("set-read-state")]
    MailMutationResult SetReadState(
        bool isRead,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("set-flag")]
    MailMutationResult SetFlag(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string flagStatus = "flagged",
        string? dueDate = null,
        string? flagRequest = null);

    [ServiceAction("set-categories")]
    MailMutationResult SetCategories(
        string? categories = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("list-categories")]
    MailCategoryListResult ListCategories();

    /// <summary>
    /// Creates a category in the mailbox's master category list - the same list <c>list-categories</c>
    /// reads and <c>set-categories</c> writes names from.
    ///
    /// <para>
    /// This is what makes <c>set-categories</c> safe. Outlook does not validate the string
    /// <c>set-categories</c> writes, so assigning a name that is not in the list quietly produces a
    /// category with no colour that the user cannot filter by. Create it first, with a colour, and
    /// the label is real.
    /// </para>
    ///
    /// <para>
    /// <b>Colour</b> is a friendly name - <c>red</c>, <c>yellow</c>, <c>darkTeal</c> and so on, exactly
    /// the names <c>list-categories</c> reports - not a raw enum ordinal. Omitting it, or naming a
    /// colour Outlook does not know, creates the category with no colour (<c>none</c>); the result
    /// reports the colour actually applied so the caller is never left guessing. A name already in the
    /// list is refused rather than duplicated.
    /// </para>
    /// </summary>
    /// <param name="name">The category name. This is how the category is addressed everywhere else, so it must not be blank.</param>
    /// <param name="color">A friendly colour name such as <c>blue</c> or <c>darkOlive</c>. Omit for no colour. An unrecognised name is treated as no colour and reported as such rather than failing.</param>
    /// <param name="shortcutKey">An optional Outlook shortcut such as <c>ctrlF2</c> through <c>ctrlF12</c>. Omit for none.</param>
    [ServiceAction("create-category", Destructive = true)]
    MailCategoryResult CreateCategory(string name, string? color = null, string? shortcutKey = null);

    /// <summary>
    /// Changes a category in the master category list: its name, its colour, its shortcut, or any
    /// combination. Addresses the category by its current name.
    ///
    /// <para>
    /// Renaming a category here does not retag the messages already carrying it - they keep the old
    /// string and lose their colour, exactly as they would in Outlook's own category manager. Prefer
    /// changing only the colour or shortcut unless the user really means to rename.
    /// </para>
    /// </summary>
    /// <param name="name">The current name of the category to change. Refused if no category by this name exists.</param>
    /// <param name="newName">A new name for the category. Omit to leave the name unchanged. Refused if another category already has this name.</param>
    /// <param name="color">A friendly colour name to apply, or <c>none</c> to clear the colour. Omit to leave the colour unchanged.</param>
    /// <param name="shortcutKey">A shortcut such as <c>ctrlF2</c>, or <c>none</c> to clear it. Omit to leave the shortcut unchanged.</param>
    [ServiceAction("update-category", Destructive = true)]
    MailCategoryResult UpdateCategory(
        string name,
        string? newName = null,
        string? color = null,
        string? shortcutKey = null);

    /// <summary>
    /// Removes a category from the master category list.
    ///
    /// <para>
    /// This is not confirmation-gated and does not touch any message: the messages already tagged
    /// with the category keep the string, they simply lose the colour and the entry in the manager.
    /// Recreate the category to restore it.
    /// </para>
    /// </summary>
    /// <param name="name">The name of the category to remove. Refused if no category by this name exists, so a caller can tell a real removal from a no-op.</param>
    [ServiceAction("delete-category", Destructive = true)]
    MailCategoryResult DeleteCategory(string name);

    [ServiceAction("list-rules")]
    MailRuleListResult ListRules(bool includeDetail = false);

    /// <summary>
    /// The reminders Outlook is holding, earliest first.
    /// </summary>
    /// <param name="maxCount">How many rows to return. The counts always describe the full set.</param>
    /// <param name="upcomingOnly">Keep only reminders that have not yet fallen due. On by default, because most reminders on a long-lived mailbox are years overdue and including them buries the ones that matter.</param>
    [ServiceAction("list-reminders")]
    MailReminderListResult ListReminders(int maxCount = 50, bool upcomingOnly = true);

    [ServiceAction("set-subject")]
    MailMutationResult SetSubject(
        string subject,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

    [ServiceAction("set-body")]
    MailMutationResult SetBody(
        string body,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string bodyFormat = "plain");

    [ServiceAction("set-recipients")]
    MailMutationResult SetRecipients(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);
}

