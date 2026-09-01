using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Mail;

public class MailCommands : IMailCommands
{
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> FolderAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["inbox"] = Outlook.OlDefaultFolders.olFolderInbox,
        ["drafts"] = Outlook.OlDefaultFolders.olFolderDrafts,
        ["sent"] = Outlook.OlDefaultFolders.olFolderSentMail,
        ["sent-mail"] = Outlook.OlDefaultFolders.olFolderSentMail,
        ["outbox"] = Outlook.OlDefaultFolders.olFolderOutbox,
        ["deleted"] = Outlook.OlDefaultFolders.olFolderDeletedItems,
        ["deleted-items"] = Outlook.OlDefaultFolders.olFolderDeletedItems,
        ["calendar"] = Outlook.OlDefaultFolders.olFolderCalendar,
        ["contacts"] = Outlook.OlDefaultFolders.olFolderContacts,
        ["tasks"] = Outlook.OlDefaultFolders.olFolderTasks,
        ["notes"] = Outlook.OlDefaultFolders.olFolderNotes,
        ["junk"] = Outlook.OlDefaultFolders.olFolderJunk
    };

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ActiveMailResult ReadActive()
        => Read(entryId: null, storeId: null, useActiveMail: true);

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public ActiveMailResult Read(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookMailRead",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new ActiveMailResult
                        {
                            Success = true,
                            HasActiveMail = false
                        };
                    }

                    return CreateActiveMailResult(mail);
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new ActiveMailResult
            {
                Success = false,
                HasActiveMail = false,
                ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                    ? $"Failed to inspect the active Outlook mail item: {ex.Message}"
                    : $"Failed to inspect the requested Outlook mail item: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailListResult List(
        string? folder = null,
        int maxCount = 25,
        bool unreadOnly = false,
        bool includeBodyPreview = false)
        => ExecuteMailList(
            "OutlookMailList",
            folder,
            query: null,
            maxCount,
            unreadOnly,
            includeBodyPreview);

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailListResult Search(
        string query,
        string? folder = null,
        int maxCount = 25,
        bool unreadOnly = false,
        bool includeBodyPreview = false)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return new MailListResult
            {
                Success = false,
                ErrorMessage = "query is required for mail.search."
            };
        }

        return ExecuteMailList(
            "OutlookMailSearch",
            folder,
            query,
            maxCount,
            unreadOnly,
            includeBodyPreview);
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult CreateMailDraft(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? subject = null,
        string? body = null,
        bool display = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookMailCreateDraft",
            (application, session) =>
            {
                object? createdItem = null;
                Outlook.MailItem? mail = null;

                try
                {
                    createdItem = application.CreateItem(Outlook.OlItemType.olMailItem);
                    mail = createdItem as Outlook.MailItem;
                    if (mail == null)
                    {
                        return new MailDraftResult
                        {
                            Success = false,
                            Saved = false,
                            Displayed = false,
                            ErrorMessage = "Outlook did not return a mail draft item."
                        };
                    }

                    if (!string.IsNullOrWhiteSpace(recipientTo))
                    {
                        mail.To = recipientTo;
                    }

                    if (!string.IsNullOrWhiteSpace(cc))
                    {
                        mail.CC = cc;
                    }

                    if (!string.IsNullOrWhiteSpace(bcc))
                    {
                        mail.BCC = bcc;
                    }

                    if (subject != null)
                    {
                        mail.Subject = subject;
                    }

                    if (body != null)
                    {
                        mail.Body = body;
                    }

                    mail.Save();

                    if (display)
                    {
                        mail.Display(false);
                    }

                    return new MailDraftResult
                    {
                        Success = true,
                        Saved = true,
                        Displayed = display,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        To = SafeGet(() => mail.To),
                        Cc = SafeGet(() => mail.CC),
                        Bcc = SafeGet(() => mail.BCC),
                        BodyLength = body?.Length ?? SafeGet(() => mail.Body)?.Length ?? 0,
                        Message = "Created Outlook draft."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                    OutlookInteropRunner.ReleaseComObject(ref createdItem);
                }
            },
            ex => new MailDraftResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = $"Failed to create an Outlook draft: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult Reply(bool display = false)
        => ExecuteDraftFromActiveMail(
            "OutlookMailReply",
            "Created Outlook reply draft.",
            display,
            sourceMail => sourceMail.Reply());

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult ReplyAll(bool display = false)
        => ExecuteDraftFromActiveMail(
            "OutlookMailReplyAll",
            "Created Outlook reply-all draft.",
            display,
            sourceMail => sourceMail.ReplyAll());

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult Forward(bool display = false)
        => ExecuteDraftFromActiveMail(
            "OutlookMailForward",
            "Created Outlook forward draft.",
            display,
            sourceMail => sourceMail.Forward());

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailSendResult Send(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookMailSend",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new MailSendResult
                        {
                            Success = false,
                            Sent = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                                ? "No active Outlook mail item is currently selected or open."
                                : "The requested Outlook mail item could not be resolved."
                        };
                    }

                    if (SafeGetBool(() => mail.Sent))
                    {
                        return new MailSendResult
                        {
                            Success = false,
                            Sent = false,
                            EntryId = SafeGet(() => mail.EntryID),
                            StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                            Subject = SafeGet(() => mail.Subject),
                            To = SafeGet(() => mail.To),
                            ErrorMessage = "The selected Outlook mail item has already been sent."
                        };
                    }

                    string? subject = SafeGet(() => mail.Subject);
                    string? recipientTo = SafeGet(() => mail.To);
                    string? resolvedEntryId = SafeGet(() => mail.EntryID);
                    string? resolvedStoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null);

                    mail.Send();

                    return new MailSendResult
                    {
                        Success = true,
                        Sent = true,
                        EntryId = resolvedEntryId,
                        StoreId = resolvedStoreId,
                        Subject = subject,
                        To = recipientTo,
                        SentOn = SafeGetDateTimeOffset(() => mail.SentOn),
                        Message = "Sent Outlook mail item."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailSendResult
            {
                Success = false,
                Sent = false,
                ErrorMessage = $"Failed to send the Outlook mail item: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult Move(
        string targetFolder,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        if (string.IsNullOrWhiteSpace(targetFolder))
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = "targetFolder is required for mail.move."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailMove",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.MailItem? movedMail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                Outlook.MAPIFolder? destinationFolder = null;
                object? movedItem = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    destinationFolder = ResolveFolder(application, session, targetFolder, ref explorer);
                    if (destinationFolder == null)
                    {
                        return new MailMutationResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(targetFolder)
                        };
                    }

                    string? sourcePath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                        ? OutlookInteropRunner.GetFolderPath(folder)
                        : null);
                    string? destinationPath = OutlookInteropRunner.GetFolderPath(destinationFolder);
                    if (!string.IsNullOrWhiteSpace(sourcePath)
                        && !string.IsNullOrWhiteSpace(destinationPath)
                        && string.Equals(sourcePath, destinationPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return new MailMutationResult
                        {
                            Success = true,
                            Moved = true,
                            EntryId = SafeGet(() => mail.EntryID),
                            StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                            Subject = SafeGet(() => mail.Subject),
                            FolderName = SafeGet(() => destinationFolder.Name),
                            FolderPath = destinationPath,
                            Read = !SafeGetBool(() => mail.UnRead),
                            Message = "The Outlook mail item is already in the requested folder."
                        };
                    }

                    movedItem = mail.Move(destinationFolder);
                    movedMail = movedItem as Outlook.MailItem;
                    if (movedMail != null)
                    {
                        movedItem = null;
                    }

                    return new MailMutationResult
                    {
                        Success = true,
                        Moved = true,
                        EntryId = SafeGet(() => movedMail?.EntryID ?? mail.EntryID),
                        StoreId = SafeGet(() => movedMail?.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => movedMail?.Subject ?? mail.Subject),
                        FolderName = SafeGet(() => destinationFolder.Name),
                        FolderPath = destinationPath,
                        Read = !SafeGetBool(() => (movedMail ?? mail).UnRead),
                        Message = "Moved Outlook mail item."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref movedMail);
                    OutlookInteropRunner.ReleaseComObject(ref movedItem);
                    OutlookInteropRunner.ReleaseComObject(ref destinationFolder);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to move the Outlook mail item: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult Delete(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookMailDelete",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    var result = new MailMutationResult
                    {
                        Success = true,
                        Deleted = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        FolderName = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.Name : null),
                        FolderPath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                            ? OutlookInteropRunner.GetFolderPath(folder)
                            : null),
                        Read = !SafeGetBool(() => mail.UnRead),
                        Message = "Deleted Outlook mail item."
                    };

                    mail.Delete();
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to delete the Outlook mail item: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult SetReadState(
        bool isRead,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookMailSetReadState",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    mail.UnRead = !isRead;
                    mail.Save();

                    return new MailMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        FolderName = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.Name : null),
                        FolderPath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                            ? OutlookInteropRunner.GetFolderPath(folder)
                            : null),
                        Read = isRead,
                        Message = isRead
                            ? "Marked Outlook mail item as read."
                            : "Marked Outlook mail item as unread."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to update the Outlook mail read state: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult SetCategories(
        string? categories = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookMailSetCategories",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    string normalizedCategories = NormalizeCategoriesForOutlook(categories);
                    mail.Categories = string.IsNullOrWhiteSpace(normalizedCategories) ? null : normalizedCategories;
                    mail.Save();

                    return new MailMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        FolderName = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.Name : null),
                        FolderPath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                            ? OutlookInteropRunner.GetFolderPath(folder)
                            : null),
                        Categories = ParseCategories(SafeGet(() => mail.Categories)),
                        Read = !SafeGetBool(() => mail.UnRead),
                        Message = string.IsNullOrWhiteSpace(normalizedCategories)
                            ? "Cleared Outlook mail categories."
                            : "Updated Outlook mail categories."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to update Outlook mail categories: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult SetSubject(
        string subject,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        if (subject == null)
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = "subject is required for mail.set-subject."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailSetSubject",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    mail.Subject = subject;
                    mail.Save();

                    return new MailMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        FolderName = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.Name : null),
                        FolderPath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                            ? OutlookInteropRunner.GetFolderPath(folder)
                            : null),
                        Categories = ParseCategories(SafeGet(() => mail.Categories)),
                        Read = !SafeGetBool(() => mail.UnRead),
                        Message = "Updated Outlook mail subject."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to update the Outlook mail subject: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult SetBody(
        string body,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        if (body == null)
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = "body is required for mail.set-body."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailSetBody",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    mail.Body = body;
                    mail.Save();

                    return new MailMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        FolderName = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.Name : null),
                        FolderPath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                            ? OutlookInteropRunner.GetFolderPath(folder)
                            : null),
                        Categories = ParseCategories(SafeGet(() => mail.Categories)),
                        Read = !SafeGetBool(() => mail.UnRead),
                        Message = "Updated Outlook mail body."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to update the Outlook mail body: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult SetRecipients(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        if (recipientTo == null && cc == null && bcc == null)
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = "At least one recipient field must be provided for mail.set-recipients."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailSetRecipients",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        entryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return CreateMailMutationNotFoundResult(entryId);
                    }

                    if (recipientTo != null)
                    {
                        mail.To = recipientTo;
                    }

                    if (cc != null)
                    {
                        mail.CC = cc;
                    }

                    if (bcc != null)
                    {
                        mail.BCC = bcc;
                    }

                    mail.Save();

                    return new MailMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        To = SafeGet(() => mail.To),
                        Cc = SafeGet(() => mail.CC),
                        Bcc = SafeGet(() => mail.BCC),
                        FolderName = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.Name : null),
                        FolderPath = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder
                            ? OutlookInteropRunner.GetFolderPath(folder)
                            : null),
                        Categories = ParseCategories(SafeGet(() => mail.Categories)),
                        Read = !SafeGetBool(() => mail.UnRead),
                        Message = "Updated Outlook mail recipients."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to update Outlook mail recipients: {ex.Message}"
            });
    }

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

    private static DateTimeOffset? SafeGetDateTimeOffset(Func<DateTime> getter)
    {
        try
        {
            DateTime value = getter();
            if (value == default)
            {
                return null;
            }

            return new DateTimeOffset(value);
        }
        catch
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static ActiveMailResult CreateActiveMailResult(Outlook.MailItem mail)
    {
        Outlook.MAPIFolder? parentFolder = null;

        try
        {
            parentFolder = mail.Parent as Outlook.MAPIFolder;

            return new ActiveMailResult
            {
                Success = true,
                HasActiveMail = true,
                EntryId = SafeGet(() => mail.EntryID),
                StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                Subject = SafeGet(() => mail.Subject),
                To = SafeGet(() => mail.To),
                Cc = SafeGet(() => mail.CC),
                Bcc = SafeGet(() => mail.BCC),
                SenderName = SafeGet(() => mail.SenderName),
                SenderEmailAddress = SafeGet(() => mail.SenderEmailAddress),
                CurrentFolderPath = OutlookInteropRunner.GetFolderPath(parentFolder),
                BodyPreview = OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body)),
                Categories = ParseCategories(SafeGet(() => mail.Categories)),
                Unread = SafeGetBool(() => mail.UnRead),
                Importance = SafeGetInt(() => (int)mail.Importance),
                AttachmentCount = SafeGetInt(() => mail.Attachments.Count),
                ReceivedTime = SafeGetDateTimeOffset(() => mail.ReceivedTime),
                SentOn = SafeGetDateTimeOffset(() => mail.SentOn)
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref parentFolder);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static MailListResult ExecuteMailList(
        string operationName,
        string? folder,
        string? query,
        int maxCount,
        bool unreadOnly,
        bool includeBodyPreview)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);
        int scanLimit = Math.Clamp(boundedMaxCount * 10, 25, 500);

        return OutlookInteropRunner.Execute(
            operationName,
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                object? items = null;

                try
                {
                    explorer = application.ActiveExplorer();
                    resolvedFolder = ResolveFolder(application, session, folder, ref explorer);
                    if (resolvedFolder == null)
                    {
                        return new MailListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = resolvedFolder.Items;
                    int totalItemCount = SafeGetInt(() => ((dynamic)items).Count);
                    TrySortItemsByReceivedTime(items);

                    var result = new MailListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        Query = query,
                        TotalItemCount = totalItemCount
                    };

                    for (int index = 1, scanned = 0;
                         index <= totalItemCount && scanned < scanLimit && result.Messages.Count < boundedMaxCount;
                         index++)
                    {
                        object? rawItem = null;
                        Outlook.MailItem? mail = null;

                        try
                        {
                            rawItem = ((dynamic)items)[index];
                            scanned++;
                            mail = rawItem as Outlook.MailItem;
                            if (mail == null)
                            {
                                continue;
                            }

                            if (unreadOnly && !SafeGetBool(() => mail.UnRead))
                            {
                                continue;
                            }

                            if (!MatchesQuery(mail, query))
                            {
                                continue;
                            }

                            result.Messages.Add(CreateMailSummary(mail, includeBodyPreview));
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref mail);
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ReturnedCount = result.Messages.Count;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new MailListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook mail items: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static MailDraftResult ExecuteDraftFromActiveMail(
        string operationName,
        string successMessage,
        bool display,
        Func<Outlook.MailItem, Outlook.MailItem> createDraft)
    {
        return OutlookInteropRunner.Execute(
            operationName,
            (application, session) =>
            {
                Outlook.MailItem? sourceMail = null;
                Outlook.MailItem? draftMail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;

                try
                {
                    inspector = application.ActiveInspector();
                    if (inspector != null)
                    {
                        currentItem = inspector.CurrentItem;
                        sourceMail = currentItem as Outlook.MailItem;
                    }

                    if (sourceMail == null)
                    {
                        explorer = application.ActiveExplorer();
                        if (explorer != null)
                        {
                            selection = explorer.Selection;
                            if (selection != null && selection.Count > 0)
                            {
                                selectedItem = selection[1];
                                sourceMail = selectedItem as Outlook.MailItem;
                            }
                        }
                    }

                    if (sourceMail == null)
                    {
                        return new MailDraftResult
                        {
                            Success = false,
                            Saved = false,
                            Displayed = false,
                            ErrorMessage = "No active Outlook mail item is currently selected or open."
                        };
                    }

                    draftMail = createDraft(sourceMail);
                    draftMail.Save();

                    if (display)
                    {
                        draftMail.Display(false);
                    }

                    return new MailDraftResult
                    {
                        Success = true,
                        Saved = true,
                        Displayed = display,
                        EntryId = SafeGet(() => draftMail.EntryID),
                        StoreId = SafeGet(() => draftMail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => draftMail.Subject),
                        To = SafeGet(() => draftMail.To),
                        Cc = SafeGet(() => draftMail.CC),
                        Bcc = SafeGet(() => draftMail.BCC),
                        BodyLength = SafeGet(() => draftMail.Body)?.Length ?? 0,
                        Message = successMessage
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref draftMail);
                    OutlookInteropRunner.ReleaseComObject(ref sourceMail);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                }
            },
            ex => new MailDraftResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = $"Failed to create a draft from the active Outlook mail item: {ex.Message}"
            });
    }

    private static Outlook.MAPIFolder? ResolveFolder(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? folder,
        ref Outlook.Explorer? explorer)
        => OutlookInteropRunner.ResolveFolder(application, session, folder, FolderAliases, ref explorer);

    private static MailMutationResult CreateMailMutationNotFoundResult(string? entryId)
        => new()
        {
            Success = false,
            ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                ? "No active Outlook mail item is currently selected or open."
                : "The requested Outlook mail item could not be resolved."
        };

    private static List<string> ParseCategories(string? categories)
        => string.IsNullOrWhiteSpace(categories)
            ? []
            :
            [
                .. categories
                    .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
            ];

    private static string NormalizeCategoriesForOutlook(string? categories)
        => string.Join(", ", ParseCategories(categories));

    private static string BuildUnknownFolderMessage(string? folder)
    {
        const string supportedFolders = "current, inbox, drafts, sent, outbox, deleted, calendar, contacts, tasks, notes, junk";
        return string.IsNullOrWhiteSpace(folder)
            ? $"Could not resolve the Outlook folder. Supported folder values: {supportedFolders}."
            : $"Unsupported Outlook folder '{folder}'. Supported folder values: {supportedFolders}.";
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void TrySortItemsByReceivedTime(object items)
    {
        try
        {
            ((dynamic)items).Sort("[ReceivedTime]", true);
        }
        catch
        {
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool MatchesQuery(Outlook.MailItem mail, string? query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return true;
        }

        string searchText = query.Trim();
        return ContainsIgnoreCase(SafeGet(() => mail.Subject), searchText)
            || ContainsIgnoreCase(SafeGet(() => mail.SenderName), searchText)
            || ContainsIgnoreCase(SafeGet(() => mail.SenderEmailAddress), searchText)
            || ContainsIgnoreCase(SafeGet(() => mail.To), searchText)
            || ContainsIgnoreCase(SafeGet(() => mail.CC), searchText)
            || ContainsIgnoreCase(OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body), maxLength: 1200), searchText);
    }

    private static bool ContainsIgnoreCase(string? value, string searchText)
        => !string.IsNullOrWhiteSpace(value)
           && value.Contains(searchText, StringComparison.OrdinalIgnoreCase);

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static MailSummaryInfo CreateMailSummary(Outlook.MailItem mail, bool includeBodyPreview)
    {
        return new MailSummaryInfo
        {
            EntryId = SafeGet(() => mail.EntryID),
            StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
            Subject = SafeGet(() => mail.Subject),
            SenderName = SafeGet(() => mail.SenderName),
            SenderEmailAddress = SafeGet(() => mail.SenderEmailAddress),
            To = SafeGet(() => mail.To),
            Cc = SafeGet(() => mail.CC),
            BodyPreview = includeBodyPreview
                ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body))
                : null,
            Categories = ParseCategories(SafeGet(() => mail.Categories)),
            Unread = SafeGetBool(() => mail.UnRead),
            IsDraft = SafeGetBool(() => !mail.Sent && SafeGet(() => mail.MessageClass)?.Contains("IPM.Note", StringComparison.OrdinalIgnoreCase) == true),
            Importance = SafeGetInt(() => (int)mail.Importance),
            AttachmentCount = SafeGetInt(() => mail.Attachments.Count),
            ReceivedTime = SafeGetDateTimeOffset(() => mail.ReceivedTime),
            SentOn = SafeGetDateTimeOffset(() => mail.SentOn)
        };
    }
}
