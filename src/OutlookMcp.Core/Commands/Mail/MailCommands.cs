using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
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
        bool includeBodyPreview = false,
        string? fromAddress = null,
        string? subjectContains = null,
        string? receivedAfter = null,
        string? receivedBefore = null,
        bool? hasAttachment = null)
        => ExecuteMailList(
            "OutlookMailList",
            folder,
            query: null,
            maxCount,
            unreadOnly,
            includeBodyPreview,
            fromAddress,
            subjectContains,
            receivedAfter,
            receivedBefore,
            hasAttachment);

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailListResult Search(
        string query,
        string? folder = null,
        int maxCount = 25,
        bool unreadOnly = false,
        bool includeBodyPreview = false,
        string? fromAddress = null,
        string? subjectContains = null,
        string? receivedAfter = null,
        string? receivedBefore = null,
        bool? hasAttachment = null)
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
            includeBodyPreview,
            fromAddress,
            subjectContains,
            receivedAfter,
            receivedBefore,
            hasAttachment);
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
    public MailDraftResult Reply(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false)
        => ExecuteDraftFromMail(
            "OutlookMailReply",
            "Created Outlook reply draft.",
            entryId,
            storeId,
            useActiveMail,
            recipientTo: null,
            cc: null,
            bcc: null,
            body,
            display,
            sourceMail => sourceMail.Reply());

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult ReplyAll(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false)
        => ExecuteDraftFromMail(
            "OutlookMailReplyAll",
            "Created Outlook reply-all draft.",
            entryId,
            storeId,
            useActiveMail,
            recipientTo: null,
            cc: null,
            bcc: null,
            body,
            display,
            sourceMail => sourceMail.ReplyAll());

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult Forward(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? body = null,
        bool display = false)
        => ExecuteDraftFromMail(
            "OutlookMailForward",
            "Created Outlook forward draft.",
            entryId,
            storeId,
            useActiveMail,
            recipientTo,
            cc,
            bcc,
            body,
            display,
            sourceMail => sourceMail.Forward());

    /// <summary>
    /// At-most-once idempotency cache for <see cref="Send"/>, keyed by caller-supplied
    /// <c>operationId</c>. If a caller retries a send with the same <c>operationId</c> (e.g.
    /// after a client-side timeout with an indeterminate outcome, see #29), the cached result
    /// from the first attempt is replayed instead of re-invoking <c>MailItem.Send()</c>, which
    /// would risk sending a duplicate message. Entries are process-lifetime (not persisted); a
    /// crash/restart loses the cache, which is an accepted tradeoff since operationId reuse is
    /// only expected within a single client session's short retry window.
    /// </summary>
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, MailSendResult> SendResultCache = new();

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailSendResult Send(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool confirm = false,
        string? operationId = null)
    {
        if (!string.IsNullOrEmpty(operationId) && SendResultCache.TryGetValue(operationId, out MailSendResult? cached))
        {
            // Replay the first attempt's outcome rather than re-sending. See #29: retrying a send
            // whose true outcome is unknown (e.g. after a timeout) must never risk duplicating it.
            return cached;
        }

        if (!confirm)
        {
            return new MailSendResult
            {
                Success = false,
                Sent = false,
                ErrorMessage = "Sending mail requires confirm=true. This is a deliberate confirmation gate " +
                               "for a destructive, one-way action (#29) -- call send again with confirm=true " +
                               "once you have verified the draft's recipients/subject/body are correct."
            };
        }

        MailSendResult result;
        try
        {
            result = OutlookInteropRunner.Execute(
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
        catch (TimeoutException ex)
        {
            // The dispatcher timed out queuing or running this operation. Since mail.Send() may
            // already have been issued to Outlook when the timeout fired, we cannot know whether
            // the message actually sent -- report Indeterminate rather than Success=false so the
            // caller does not treat this as "definitely not sent" and blindly retry (which risks
            // a duplicate send). See #29.
            result = new MailSendResult
            {
                Success = false,
                Sent = false,
                Indeterminate = true,
                ErrorMessage = $"Send timed out; the outcome is unknown (the message may have been sent). " +
                               $"Re-check via mail.read before retrying -- do not blindly resend. {ex.Message}"
            };
        }

        if (!string.IsNullOrEmpty(operationId))
        {
            SendResultCache.TryAdd(operationId, result);
        }

        return result;
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

    /// <summary>
    /// Safely reads a string-valued Outlook property. If the read fails because the Outlook
    /// Object Model Guard blocked it (see #30), the failure is recorded in
    /// <paramref name="accessDenied"/> by <paramref name="propertyName"/> instead of being
    /// silently indistinguishable from "property not present" -- see Rule 22.
    /// </summary>
    private static string? SafeGet(Func<string?> getter, string propertyName, List<string> accessDenied)
    {
        try
        {
            return getter();
        }
        catch (COMException ex) when (OutlookInteropRunner.IsObjectModelGuardDenial(ex))
        {
            accessDenied.Add(propertyName);
            return null;
        }
        catch
        {
            return null;
        }
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
        var accessDenied = new List<string>();

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
                To = SafeGet(() => mail.To, nameof(ActiveMailResult.To), accessDenied),
                Cc = SafeGet(() => mail.CC, nameof(ActiveMailResult.Cc), accessDenied),
                Bcc = SafeGet(() => mail.BCC, nameof(ActiveMailResult.Bcc), accessDenied),
                SenderName = SafeGet(() => mail.SenderName),
                SenderEmailAddress = SafeGet(() => mail.SenderEmailAddress, nameof(ActiveMailResult.SenderEmailAddress), accessDenied),
                CurrentFolderPath = OutlookInteropRunner.GetFolderPath(parentFolder),
                BodyPreview = OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body)),
                Categories = ParseCategories(SafeGet(() => mail.Categories)),
                AccessDenied = accessDenied.Count > 0 ? accessDenied : null,
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
        bool includeBodyPreview,
        string? fromAddress = null,
        string? subjectContains = null,
        string? receivedAfter = null,
        string? receivedBefore = null,
        bool? hasAttachment = null)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);
        // Safety net only (not a "found" cutoff): protects against pathologically slow scans of
        // huge folders. Previously this doubled as an undocumented hard cap on how far back
        // mail.list/search could ever see -- anything beyond it was silently invisible, which an
        // LLM reads as "no such mail exists" (#27). Restrict() below now does the actual
        // filtering at the Outlook/MAPI layer instead of via client-side scanning, so this cap is
        // only hit for very large folders/queries, and MailListResult.Truncated makes that
        // explicit instead of silent.
        const int ScanSafetyLimit = 5000;

        if (!TryParseFilterDate(receivedAfter, nameof(receivedAfter), out DateTimeOffset? parsedAfter, out string? dateError)
            || !TryParseFilterDate(receivedBefore, nameof(receivedBefore), out DateTimeOffset? parsedBefore, out dateError))
        {
            return new MailListResult { Success = false, ErrorMessage = dateError };
        }

        if (parsedAfter.HasValue && parsedBefore.HasValue && parsedBefore.Value < parsedAfter.Value)
        {
            return new MailListResult
            {
                Success = false,
                ErrorMessage = "receivedBefore must be greater than or equal to receivedAfter."
            };
        }

        string? restrictFilter = MailRestrictFilter.Build(
            unreadOnly,
            fromAddress,
            subjectContains,
            parsedAfter,
            parsedBefore,
            hasAttachment);

        return OutlookInteropRunner.Execute(
            operationName,
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Items? items = null;
                Outlook.Items? restrictedItems = null;

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
                    int totalItemCount = SafeGetInt(() => items.Count);
                    TrySortItemsByReceivedTime(items);

                    // Push the structured predicates down to Outlook via Items.Restrict (DASL)
                    // instead of hydrating every item and checking client-side -- this is both
                    // faster (Restrict returns a pre-filtered rowset) and correct (a match far back
                    // in a large folder is still found, since Restrict does not stop at any
                    // client-side scan cap). See #27.
                    Outlook.Items itemsToScan = items;
                    int scanCount = totalItemCount;
                    if (restrictFilter != null)
                    {
                        try
                        {
                            restrictedItems = items.Restrict(restrictFilter);
                            itemsToScan = restrictedItems;
                            scanCount = SafeGetInt(() => restrictedItems.Count);
                        }
                        catch (COMException)
                        {
                            // Fall back to the unfiltered folder if Restrict is unavailable for
                            // this folder/store type. The client-side checks below are applied
                            // unconditionally, so the result set is identical either way - only
                            // slower, and bounded by ScanSafetyLimit.
                            itemsToScan = items;
                            scanCount = totalItemCount;
                        }
                    }

                    var result = new MailListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        Query = query,
                        TotalItemCount = totalItemCount
                    };

                    int scanned = 0;
                    for (int index = 1;
                         index <= scanCount && scanned < ScanSafetyLimit && result.Messages.Count < boundedMaxCount;
                         index++)
                    {
                        object? rawItem = null;
                        Outlook.MailItem? mail = null;

                        try
                        {
                            rawItem = itemsToScan[index];
                            scanned++;
                            mail = rawItem as Outlook.MailItem;
                            if (mail == null)
                            {
                                continue;
                            }

                            // Applied even when Restrict succeeded. The DASL filter is deliberately
                            // over-inclusive -- it drops any predicate it cannot express exactly
                            // (see MailRestrictFilter) -- so this is what makes the result exact.
                            if (!MatchesStructuredFilters(
                                    mail, unreadOnly, fromAddress, subjectContains,
                                    parsedAfter, parsedBefore, hasAttachment))
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
                    result.ScannedCount = scanned;
                    // Truncated: there was more to look at than we actually evaluated, whether
                    // because the result cap (maxCount) was hit first or the safety limit was --
                    // either way, "no more results" must not be inferred from this response alone.
                    result.Truncated = scanned < scanCount;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref restrictedItems);
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
    private static MailDraftResult ExecuteDraftFromMail(
        string operationName,
        string successMessage,
        string? entryId,
        string? storeId,
        bool useActiveMail,
        string? recipientTo,
        string? cc,
        string? bcc,
        string? body,
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
                object? resolvedItem = null;

                try
                {
                    sourceMail = OutlookInteropRunner.ResolveMailItem(
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

                    if (sourceMail == null)
                    {
                        return new MailDraftResult
                        {
                            Success = false,
                            Saved = false,
                            Displayed = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                                // Headless targeting: no active Outlook window/selection is
                                // required when entryId is supplied. See #36.
                                ? "No active Outlook mail item is currently selected or open. " +
                                  "Pass entryId (e.g. from mail.search/mail.list) to target a specific message headlessly."
                                : "The requested Outlook mail item could not be resolved."
                        };
                    }

                    draftMail = createDraft(sourceMail);

                    if (!string.IsNullOrEmpty(recipientTo))
                    {
                        draftMail.To = recipientTo;
                    }

                    if (!string.IsNullOrEmpty(cc))
                    {
                        draftMail.CC = cc;
                    }

                    if (!string.IsNullOrEmpty(bcc))
                    {
                        draftMail.BCC = bcc;
                    }

                    if (body != null)
                    {
                        // Reply()/ReplyAll()/Forward() pre-populate Body with the quoted original
                        // message; prepending the caller's text keeps that quoted context instead
                        // of destroying it, matching how a person would type above a quoted reply.
                        string existingBody = SafeGet(() => draftMail.Body) ?? string.Empty;
                        draftMail.Body = body + Environment.NewLine + Environment.NewLine + existingBody;
                    }

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
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref sourceMail);
                }
            },
            ex => new MailDraftResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = $"Failed to create a draft from the Outlook mail item: {ex.Message}"
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
    private static void TrySortItemsByReceivedTime(Outlook.Items items)
    {
        try
        {
            items.Sort("[ReceivedTime]", true);
        }
        catch
        {
        }
    }

    /// <summary>
    /// Applies the structured predicates client-side. This runs even when <c>Restrict</c> succeeded:
    /// the DASL filter is deliberately over-inclusive (it omits any predicate it cannot express
    /// exactly, such as a value containing a <c>LIKE</c> wildcard), so this is the step that makes
    /// the result exact. Running it unconditionally also means the Restrict path and the fallback
    /// path return identical results, differing only in speed.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool MatchesStructuredFilters(
        Outlook.MailItem mail,
        bool unreadOnly,
        string? fromAddress,
        string? subjectContains,
        DateTimeOffset? receivedAfter,
        DateTimeOffset? receivedBefore,
        bool? hasAttachment)
    {
        if (unreadOnly && !SafeGetBool(() => mail.UnRead))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(subjectContains)
            && !ContainsIgnoreCase(SafeGet(() => mail.Subject), subjectContains.Trim()))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(fromAddress))
        {
            string needle = fromAddress.Trim();
            if (!ContainsIgnoreCase(SafeGet(() => mail.SenderEmailAddress), needle)
                && !ContainsIgnoreCase(SafeGet(() => mail.SenderName), needle))
            {
                return false;
            }
        }

        if (hasAttachment.HasValue && SafeGetInt(() => mail.Attachments.Count) > 0 != hasAttachment.Value)
        {
            return false;
        }

        if (receivedAfter.HasValue || receivedBefore.HasValue)
        {
            DateTimeOffset? received = SafeGetDateTimeOffset(() => mail.ReceivedTime);
            if (received == null)
            {
                // A message whose ReceivedTime cannot be read (an unsent draft, for instance)
                // cannot be shown to satisfy a date window, so it is excluded rather than guessed at.
                return false;
            }

            if (receivedAfter.HasValue && received.Value < receivedAfter.Value)
            {
                return false;
            }

            if (receivedBefore.HasValue && received.Value > receivedBefore.Value)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Parses a caller-supplied filter date, mirroring the ISO-8601 handling the calendar actions
    /// already use so that a date means the same thing across the whole tool surface.
    /// </summary>
    private static bool TryParseFilterDate(
        string? value,
        string parameterName,
        out DateTimeOffset? parsed,
        out string? errorMessage)
    {
        parsed = null;
        errorMessage = null;

        if (string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        if (!DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTimeOffset result))
        {
            errorMessage = $"{parameterName} must be a valid ISO date/time value (for example 2024-03-07 or 2024-03-07T14:30).";
            return false;
        }

        parsed = result;
        return true;
    }

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
        var accessDenied = new List<string>();

        var summary = new MailSummaryInfo
        {
            EntryId = SafeGet(() => mail.EntryID),
            StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
            Subject = SafeGet(() => mail.Subject),
            SenderName = SafeGet(() => mail.SenderName),
            SenderEmailAddress = SafeGet(() => mail.SenderEmailAddress, nameof(MailSummaryInfo.SenderEmailAddress), accessDenied),
            To = SafeGet(() => mail.To, nameof(MailSummaryInfo.To), accessDenied),
            Cc = SafeGet(() => mail.CC, nameof(MailSummaryInfo.Cc), accessDenied),
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

        summary.AccessDenied = accessDenied.Count > 0 ? accessDenied : null;
        return summary;
    }
}
