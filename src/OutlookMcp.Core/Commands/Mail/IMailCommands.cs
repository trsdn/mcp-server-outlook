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
    + "filters. It is exhaustive rather than indexed, so it will not miss a term buried deep in a long message, but "
    + "reaching the body means opening every candidate item; Outlook cannot filter on body text server-side, so pair "
    + "query with structured filters in large folders to cut how many items must be opened. "
    + "list and search page with an opaque cursor: when a response has hasMore true, pass its nextCursor back as "
    + "cursor on an otherwise identical call to get the next page, and keep going until hasMore is false. Never treat "
    + "a truncated or empty page as proof that no such mail exists while hasMore is true. A cursor only continues the "
    + "exact folder, query and filters that produced it - changing any of them requires restarting without a cursor - "
    + "but maxCount may be changed freely mid-walk. "
    + "Use get-conversation to retrieve a whole mail thread in one call - every message in the "
    + "conversation, oldest first, spanning folders, so a reply in Sent Items comes back alongside "
    + "the original in the Inbox. Prefer it over guessing which search results belong together by "
    + "subject, which is unreliable and cannot see replies filed in other folders. read, list and "
    + "search all report conversationId, so a thread can be reached from any of them. "
    + "Use create-draft to create and save a new Outlook draft with optional recipients, subject, and body text. "
    + "Use reply, reply-all, and forward to create saved draft responses targeting an explicit mail item by entry id/store id, "
    + "or fall back to the active mail item; works headlessly with no Outlook window focused when entryId is supplied. "
    + "forward accepts recipients since a forwarded message otherwise has nobody to send to. All three accept an optional "
    + "body to prepend above the quoted original message. "
    + "Use set-subject, set-body, and set-recipients to edit an existing draft before sending. "
    + "Use send to send a saved draft by entry id or the current active draft explicitly. Send requires confirm=true "
    + "(it is refused otherwise) and accepts an optional operationId so a retried call with the same operationId after "
    + "a timeout or crash is answered from a cached result instead of risking a duplicate send. "
    + "Set display=true on draft-producing actions to show the draft inspector after saving.")]
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
        string? cursor = null);

    [ServiceAction("get-conversation", Destructive = false)]
    MailConversationResult GetConversation(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        int maxCount = 50,
        bool includeBodyPreview = false);

    [ServiceAction("create-draft")]
    MailDraftResult CreateMailDraft(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? subject = null,
        string? body = null,
        bool display = false);

    [ServiceAction("reply")]
    MailDraftResult Reply(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false);

    [ServiceAction("reply-all")]
    MailDraftResult ReplyAll(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false);

    [ServiceAction("forward")]
    MailDraftResult Forward(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? body = null,
        bool display = false);

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

    [ServiceAction("set-categories")]
    MailMutationResult SetCategories(
        string? categories = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

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
        bool useActiveMail = true);

    [ServiceAction("set-recipients")]
    MailMutationResult SetRecipients(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);
}
