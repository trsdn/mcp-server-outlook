using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Mail;

[ServiceCategory("mail")]
[NoSession]
[McpTool("mail", Title = "Outlook Mail Operations", Destructive = true, Category = "mail",
    Description = "Read the currently active Outlook mail item, list and search mailbox content, and create or send Outlook draft emails without opening a persistent session. "
    + "Use read-active to inspect the currently selected or opened mail item. "
    + "Use read to inspect an explicit Outlook mail item by entry id/store id or fall back to the active mail item. "
    + "Use list and search to inspect the current folder or a default Outlook folder role such as inbox or drafts. "
    + "Use create-draft to create and save a new Outlook draft with optional recipients, subject, and body text. "
    + "Use reply, reply-all, and forward to create saved draft responses from the active mail item without sending them. "
    + "Use send to send a saved draft by entry id or the current active draft explicitly. "
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
        bool includeBodyPreview = false);

    [ServiceAction("search", Destructive = false)]
    MailListResult Search(
        string query,
        string? folder = null,
        int maxCount = 25,
        bool unreadOnly = false,
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
    MailDraftResult Reply(bool display = false);

    [ServiceAction("reply-all")]
    MailDraftResult ReplyAll(bool display = false);

    [ServiceAction("forward")]
    MailDraftResult Forward(bool display = false);

    [ServiceAction("send")]
    MailSendResult Send(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true);

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
}
