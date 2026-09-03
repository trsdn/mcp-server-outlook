using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Mail;

public class MailCommands : IMailCommands
{
    /// <summary>
    /// Upper bound on how many entries of one conversation are enumerated. A safety net for a
    /// pathological thread, not a paging cap - <see cref="MailConversationResult.TotalItemCount"/>
    /// reports what was found so hitting it is never silent.
    /// </summary>
    private const int ThreadSafetyLimit = 500;

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
        bool? hasAttachment = null,
        bool flaggedOnly = false,
        string? cursor = null)
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
            hasAttachment,
            flaggedOnly,
            cursor);

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
        bool? hasAttachment = null,
        bool flaggedOnly = false,
        string? cursor = null,
        string? searchMode = null)
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
            hasAttachment,
            flaggedOnly,
            cursor,
            searchMode);
    }

    /// <summary>
    /// Returns every message in one mail thread, oldest first, across folders (#39).
    ///
    /// <para>
    /// The thread is enumerated from Outlook's own conversation view (<c>Conversation.GetTable</c>)
    /// rather than by matching subjects, which is what makes it correct across folders: a reply
    /// filed in Sent Items and the original in the Inbox are one conversation to Outlook and no
    /// folder-scoped listing can ever assemble them.
    /// </para>
    ///
    /// <para>
    /// Each entry is then opened to read its sender, timestamp and folder. That is one item open per
    /// thread message - threads are small, and the alternative (projecting columns off the table) is
    /// unreliable for drafts and non-mail entries, which carry different properties. The scan is
    /// bounded by <see cref="ThreadSafetyLimit"/> regardless.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailConversationResult GetConversation(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        int maxCount = 50,
        bool includeBodyPreview = false)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);

        return OutlookInteropRunner.Execute(
            "OutlookMailGetConversation",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;
                Outlook.Conversation? conversation = null;
                Outlook.Table? table = null;

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
                        return new MailConversationResult
                        {
                            Success = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                                ? "No active Outlook mail item is available to read a conversation from."
                                : "The requested Outlook mail item could not be resolved, so its conversation cannot be read."
                        };
                    }

                    string? conversationId = SafeGet(() => mail.ConversationID);
                    string? conversationTopic = SafeGet(() => mail.ConversationTopic);
                    string? itemStoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null);

                    conversation = SafeGetConversation(mail);
                    if (conversation == null)
                    {
                        // Deliberately a failure, not an empty success. "This message has no replies"
                        // and "this store cannot tell you whether it has replies" are different
                        // answers, and reporting the second as the first is exactly the confidently
                        // wrong answer this surface must never give.
                        return new MailConversationResult
                        {
                            Success = false,
                            ConversationSupported = false,
                            ConversationId = conversationId,
                            ConversationTopic = conversationTopic,
                            ErrorMessage = "This message's store does not provide a conversation view, so its thread cannot be assembled. "
                                + "Fall back to mail.search on the conversation topic, and treat the result as a guess rather than the thread."
                        };
                    }

                    table = conversation.GetTable();
                    List<string> threadEntryIds = ReadThreadEntryIds(table);

                    var messages = new List<MailSummaryInfo>(threadEntryIds.Count);
                    int skipped = 0;

                    foreach (string threadEntryId in threadEntryIds)
                    {
                        object? item = null;
                        Outlook.MailItem? threadMail = null;
                        Outlook.MAPIFolder? parent = null;

                        try
                        {
                            item = session.GetItemFromID(
                                threadEntryId,
                                string.IsNullOrWhiteSpace(itemStoreId) ? Type.Missing : itemStoreId);
                            threadMail = item as Outlook.MailItem;

                            if (threadMail == null)
                            {
                                // A meeting request or delivery report filed into the same
                                // conversation. Counted rather than dropped silently, so a caller
                                // can see why the totals differ.
                                skipped++;
                                continue;
                            }

                            MailSummaryInfo summary = CreateMailSummary(threadMail, includeBodyPreview);
                            parent = threadMail.Parent as Outlook.MAPIFolder;
                            summary.FolderPath = OutlookInteropRunner.GetFolderPath(parent);
                            messages.Add(summary);
                        }
                        catch (COMException)
                        {
                            // An entry the conversation still lists but the store can no longer
                            // return - deleted mid-read, or in a store this profile cannot open.
                            skipped++;
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref parent);
                            OutlookInteropRunner.ReleaseComObject(ref threadMail);
                            OutlookInteropRunner.ReleaseComObject(ref item);
                        }
                    }

                    // Reading order. A thread returned newest-first, or in store order, is not a
                    // thread a caller can read. Items with no timestamp at all sort last rather than
                    // vanishing.
                    List<MailSummaryInfo> ordered = [.. messages
                        .OrderBy(m => m.ReceivedTime ?? m.SentOn ?? DateTimeOffset.MaxValue)];

                    bool truncated = ordered.Count > boundedMaxCount;
                    List<MailSummaryInfo> page = truncated
                        ? [.. ordered.Take(boundedMaxCount)]
                        : ordered;

                    return new MailConversationResult
                    {
                        Success = true,
                        ConversationSupported = true,
                        ConversationId = conversationId,
                        ConversationTopic = conversationTopic,
                        TotalItemCount = threadEntryIds.Count,
                        ReturnedCount = page.Count,
                        SkippedItemCount = skipped,
                        Truncated = truncated,
                        SortedBy = "receivedTime",
                        SortDirection = "ascending",
                        Messages = page
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref table);
                    OutlookInteropRunner.ReleaseComObject(ref conversation);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new MailConversationResult
            {
                Success = false,
                ErrorMessage = $"Failed to read the Outlook conversation: {ex.Message}"
            });
    }

    /// <summary>
    /// Accepts, declines or tentatively accepts a meeting invitation.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Answering and notifying are separate. <c>Respond</c> updates the caller's own calendar
    /// immediately; the response item it hands back is only mailed when <paramref name="sendResponse"/>
    /// is set. That is Outlook's own "do not send a response" behaviour, and making it the default
    /// keeps the irreversible half - mail to a real organiser - opt-in.
    /// </para>
    /// <para>
    /// A cancellation or a response is not an invitation, and neither is ordinary mail. All three
    /// turn up in a listing looking like something you could answer, so each is refused by name
    /// rather than failing obscurely inside Outlook.
    /// </para>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MeetingResponseResult RespondToMeeting(
        string? entryId = null,
        string? storeId = null,
        string response = "accept",
        bool sendResponse = false,
        string? responseText = null,
        bool useActiveMail = false)
    {
        if (!TryParseMeetingResponse(response, out Outlook.OlMeetingResponse parsedResponse, out string normalizedResponse))
        {
            return new MeetingResponseResult
            {
                Success = false,
                ResponseSent = false,
                ErrorMessage = $"'{response}' is not a valid response for mail.respond-to-meeting. "
                    + "Use accept, decline or tentative."
            };
        }

        if (string.IsNullOrWhiteSpace(entryId) && !useActiveMail)
        {
            return new MeetingResponseResult
            {
                Success = false,
                ResponseSent = false,
                Response = normalizedResponse,
                ErrorMessage = "entryId is required for mail.respond-to-meeting, or set useActiveMail to answer the "
                    + "invitation currently open or selected in Outlook."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailRespondToMeeting",
            (application, session) =>
            {
                object? resolvedItem = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                object? currentItem = null;
                Outlook.MeetingItem? invitation = null;
                Outlook.AppointmentItem? appointment = null;
                object? responseItem = null;
                Outlook.MeetingItem? responseMeeting = null;

                try
                {
                    if (!string.IsNullOrWhiteSpace(entryId))
                    {
                        resolvedItem = session.GetItemFromID(
                            entryId,
                            string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
                    }
                    else
                    {
                        inspector = application.ActiveInspector();
                        currentItem = inspector?.CurrentItem;
                        resolvedItem = currentItem;

                        if (resolvedItem is not Outlook.MeetingItem)
                        {
                            explorer = application.ActiveExplorer();
                            selection = explorer?.Selection;

                            if (selection != null && selection.Count > 0)
                            {
                                resolvedItem = selection[1];
                            }
                        }
                    }

                    if (resolvedItem == null)
                    {
                        return new MeetingResponseResult
                        {
                            Success = false,
                            ResponseSent = false,
                            Response = normalizedResponse,
                            ErrorMessage = string.IsNullOrWhiteSpace(entryId)
                                ? "No Outlook item is open or selected, so there is no invitation to answer."
                                : "The requested Outlook item could not be resolved, so there is no invitation to answer."
                        };
                    }

                    invitation = resolvedItem as Outlook.MeetingItem;

                    if (invitation == null)
                    {
                        return new MeetingResponseResult
                        {
                            Success = false,
                            ResponseSent = false,
                            Response = normalizedResponse,
                            ErrorMessage = "That Outlook item is not a meeting invitation, so it cannot be accepted or declined. "
                                + "A listing's itemType says which items are invitations ('meetingRequest')."
                        };
                    }

                    string? messageClass = SafeGet(() => invitation.MessageClass);
                    string? kind = ClassifyMeetingItem(messageClass);

                    if (kind != "meetingRequest")
                    {
                        return new MeetingResponseResult
                        {
                            Success = false,
                            ResponseSent = false,
                            Response = normalizedResponse,
                            Subject = SafeGet(() => invitation.Subject),
                            ErrorMessage = kind switch
                            {
                                "meetingCancellation" => "That item is a meeting cancellation, not an invitation - the meeting is already off, "
                                    + "so there is nothing to accept or decline.",
                                "meetingResponse" => "That item is somebody else's response to a meeting you organised, not an invitation to you, "
                                    + "so it cannot be accepted or declined.",
                                _ => $"That item is a scheduling message of class '{messageClass}', not a meeting invitation, "
                                    + "so it cannot be accepted or declined."
                            }
                        };
                    }

                    // false: read the appointment as it stands rather than adding a copy to the
                    // calendar before the caller has actually answered.
                    appointment = invitation.GetAssociatedAppointment(false);

                    if (appointment == null)
                    {
                        return new MeetingResponseResult
                        {
                            Success = false,
                            ResponseSent = false,
                            Response = normalizedResponse,
                            Subject = SafeGet(() => invitation.Subject),
                            ErrorMessage = "Outlook could not open the appointment behind this invitation, so it cannot be answered."
                        };
                    }

                    responseItem = appointment.Respond(parsedResponse, true, false);
                    responseMeeting = responseItem as Outlook.MeetingItem;

                    bool sent = false;

                    if (sendResponse)
                    {
                        if (responseMeeting == null)
                        {
                            return new MeetingResponseResult
                            {
                                Success = false,
                                ResponseSent = false,
                                Response = normalizedResponse,
                                Subject = SafeGet(() => appointment.Subject),
                                ErrorMessage = "Your calendar was updated, but Outlook did not produce a response message, "
                                    + "so the organiser has not been told."
                            };
                        }

                        if (!string.IsNullOrWhiteSpace(responseText))
                        {
                            responseMeeting.Body = responseText;
                        }

                        responseMeeting.Send();
                        sent = true;
                    }

                    return new MeetingResponseResult
                    {
                        Success = true,
                        ResponseSent = sent,
                        Response = normalizedResponse,
                        EntryId = SafeGet(() => appointment.EntryID),
                        StoreId = SafeGet(() => appointment.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => appointment.Subject),
                        Message = sent
                            ? $"Responded '{normalizedResponse}' and told the organiser."
                            : $"Responded '{normalizedResponse}' in your own calendar. The organiser has not been told - "
                                + "pass sendResponse to notify them."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref responseMeeting);
                    OutlookInteropRunner.ReleaseComObject(ref responseItem);
                    OutlookInteropRunner.ReleaseComObject(ref appointment);
                    OutlookInteropRunner.ReleaseComObject(ref invitation);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                }
            },
            ex => new MeetingResponseResult
            {
                Success = false,
                ResponseSent = false,
                Response = normalizedResponse,
                ErrorMessage = $"Failed to respond to the Outlook meeting invitation: {ex.Message}"
            });
    }

    private static bool TryParseMeetingResponse(
        string? response,
        out Outlook.OlMeetingResponse parsed,
        out string normalized)
    {
        switch (response?.Trim().ToLowerInvariant())
        {
            case "accept":
            case "accepted":
                parsed = Outlook.OlMeetingResponse.olMeetingAccepted;
                normalized = "accept";
                return true;
            case "decline":
            case "declined":
                parsed = Outlook.OlMeetingResponse.olMeetingDeclined;
                normalized = "decline";
                return true;
            case "tentative":
            case "tentatively":
                parsed = Outlook.OlMeetingResponse.olMeetingTentative;
                normalized = "tentative";
                return true;
            default:
                parsed = Outlook.OlMeetingResponse.olMeetingDeclined;
                normalized = response?.Trim() ?? string.Empty;
                return false;
        }
    }

    /// <summary>
    /// Pulls the entry ids out of a conversation table. Only <c>EntryID</c> is requested: the
    /// remaining default columns are read for every row and never used, and their meaning varies by
    /// item type.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static List<string> ReadThreadEntryIds(Outlook.Table table)
    {
        var entryIds = new List<string>();

        table.Columns.RemoveAll();
        table.Columns.Add("EntryID");

        while (!table.EndOfTable && entryIds.Count < ThreadSafetyLimit)
        {
            Outlook.Row? row = null;

            try
            {
                row = table.GetNextRow();
                if (row?["EntryID"] is string value && !string.IsNullOrWhiteSpace(value))
                {
                    entryIds.Add(value);
                }
            }
            finally
            {
                OutlookInteropRunner.ReleaseComObject(ref row);
            }
        }

        return entryIds;
    }

    /// <summary>
    /// <c>MailItem.GetConversation</c> returns null on stores without conversation view, and throws
    /// on some of them instead. Both mean the same thing to a caller, so both are normalised to null
    /// here and reported explicitly by the caller rather than surfacing as an opaque COM failure.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Conversation? SafeGetConversation(Outlook.MailItem mail)
    {
        try
        {
            return mail.GetConversation();
        }
        catch (COMException)
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult CreateMailDraft(
        string? recipientTo = null,
        string? cc = null,
        string? bcc = null,
        string? subject = null,
        string? body = null,
        bool display = false,
        string bodyFormat = "plain")
    {
        if (!TryParseBodyFormat(bodyFormat, out bool asHtml, out string? formatError))
        {
            return new MailDraftResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = formatError
            };
        }

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
                        if (asHtml)
                        {
                            mail.HTMLBody = body;
                        }
                        else
                        {
                            mail.Body = body;
                        }
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
        bool display = false,
        string bodyFormat = "plain")
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
            bodyFormat,
            sourceMail => sourceMail.Reply());

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailDraftResult ReplyAll(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string? body = null,
        bool display = false,
        string bodyFormat = "plain")
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
            bodyFormat,
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
        bool display = false,
        string bodyFormat = "plain")
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
            bodyFormat,
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

    /// <summary>
    /// Sets, completes or clears a message's follow-up flag (#15).
    ///
    /// <para>
    /// <c>complete</c> is deliberately distinct from <c>none</c>. Marking something done and never
    /// having flagged it look the same in a naive implementation, and they are not the same answer to
    /// "what still needs attention" - one is handled, the other was never raised.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public MailMutationResult SetFlag(
        string? entryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        string flagStatus = "flagged",
        string? dueDate = null,
        string? flagRequest = null)
    {
        string status = string.IsNullOrWhiteSpace(flagStatus) ? "flagged" : flagStatus.Trim();

        if (!status.Equals("flagged", StringComparison.OrdinalIgnoreCase)
            && !status.Equals("complete", StringComparison.OrdinalIgnoreCase)
            && !status.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = $"flagStatus '{flagStatus}' is not supported. Use 'flagged', 'complete' or 'none'."
            };
        }

        if (!TryParseFilterDate(dueDate, nameof(dueDate), out DateTimeOffset? parsedDue, out string? dateError))
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = dateError
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookMailSetFlag",
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

                    if (status.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // ClearTaskFlag() reports success on a draft and leaves FlagStatus untouched
                        // - measured, not assumed - so the state is assigned directly instead.
                        // Clearing the status alone also leaves the old task dates behind, which
                        // would surface as a due date on an unflagged message, so they are reset to
                        // the same far-future sentinel Outlook uses for "never set".
                        mail.FlagStatus = Outlook.OlFlagStatus.olNoFlag;
                        TrySetTaskDate(d => mail.TaskDueDate = d, NoTaskDate);
                        TrySetTaskDate(d => mail.TaskStartDate = d, NoTaskDate);
                    }
                    else if (status.Equals("complete", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete;
                        }
                        catch (Exception ex) when (ex is NotImplementedException or COMException)
                        {
                            return new MailMutationResult
                            {
                                Success = false,
                                ErrorMessage = "Outlook will not mark a draft as complete - a follow-up can only be "
                                    + "completed on a message that has been sent or received. Send the message first, "
                                    + "or use flagStatus 'none' to clear the flag."
                            };
                        }
                    }
                    else
                    {
                        // MarkAsTask is the documented way to raise a flag and is what makes the item
                        // appear under follow-up, but Outlook refuses it on drafts ("MarkAsTask is
                        // only valid on items that have been sent or received"). It signals that as
                        // E_NOTIMPL, which the interop assembly surfaces as NotImplementedException
                        // rather than COMException - measured, not assumed. Assigning the status
                        // directly does work on a draft, so that is the fallback rather than failing.
                        try
                        {
                            mail.MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);
                        }
                        catch (Exception ex) when (ex is NotImplementedException or COMException)
                        {
                            mail.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;
                        }

                        mail.FlagRequest = string.IsNullOrWhiteSpace(flagRequest) ? "Follow up" : flagRequest;

                        if (parsedDue.HasValue)
                        {
                            DateTime due = parsedDue.Value.LocalDateTime.Date;
                            TrySetTaskDate(d => mail.TaskStartDate = d, due);
                            TrySetTaskDate(d => mail.TaskDueDate = d, due);
                        }
                    }

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
                        Message = status.Equals("none", StringComparison.OrdinalIgnoreCase)
                            ? "Cleared the Outlook follow-up flag."
                            : $"Set the Outlook follow-up flag to {status.ToLowerInvariant()}."
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
                ErrorMessage = $"Failed to update the Outlook follow-up flag: {ex.Message}"
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
        bool useActiveMail = true,
        string bodyFormat = "plain")
    {
        if (body == null)
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = "body is required for mail.set-body."
            };
        }

        if (!TryParseBodyFormat(bodyFormat, out bool asHtml, out string? formatError))
        {
            return new MailMutationResult
            {
                Success = false,
                ErrorMessage = formatError
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

                    if (asHtml)
                    {
                        mail.HTMLBody = body;
                    }
                    else
                    {
                        mail.Body = body;
                    }

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
                ConversationId = SafeGet(() => mail.ConversationID),
                ConversationTopic = SafeGet(() => mail.ConversationTopic),
                BodyPreview = OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body)),
                Categories = ParseCategories(SafeGet(() => mail.Categories)),
                FlagStatus = MapFlagStatus(SafeGetInt(() => (int)mail.FlagStatus)),
                FlagRequest = NullIfBlank(SafeGet(() => mail.FlagRequest)),
                FlagDueDate = NormalizeTaskDate(SafeGetDateTimeOffset(() => mail.TaskDueDate)),
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
        bool? hasAttachment = null,
        bool flaggedOnly = false,
        string? cursor = null,
        string? searchMode = null)
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

        // An unrecognised mode is refused rather than defaulted. Falling back silently would hand
        // the caller substring semantics while they believed they had asked for the index, and the
        // difference is only visible in results they never see.
        if (!TryParseSearchMode(searchMode, out bool useContentIndex, out string? modeError))
        {
            return new MailListResult { Success = false, ErrorMessage = modeError };
        }

        // The content index answers the free-text question itself, so the query is pushed down
        // instead of being checked client-side afterwards.
        string? pushedDownQuery = useContentIndex ? query : null;

        string? restrictFilter = MailRestrictFilter.Build(
            unreadOnly,
            fromAddress,
            subjectContains,
            parsedAfter,
            parsedBefore,
            hasAttachment,
            flaggedOnly,
            pushedDownQuery);

        // A cursor is bound to the exact query that minted it (#43). maxCount is deliberately not
        // part of the fingerprint: changing page size part-way through a walk is legitimate.
        string fingerprint = MailPageCursor.BuildFingerprint(
            folder, query, unreadOnly, fromAddress, subjectContains,
            receivedAfter, receivedBefore, hasAttachment, flaggedOnly, searchMode);

        MailPageCursor? page = null;
        if (!string.IsNullOrWhiteSpace(cursor))
        {
            if (!MailPageCursor.TryDecode(cursor, fingerprint, out page, out string? cursorError))
            {
                // Deliberately a hard failure rather than a silent restart. Quietly returning page
                // one would make a caller looping on hasMore never terminate, and a caller checking
                // for anything past the first page conclude there was nothing there.
                return new MailListResult { Success = false, ErrorMessage = cursorError };
            }

            // Narrow the server-side filter to the cursor boundary as well, so continuing a walk
            // does not re-scan every page already visited. The DASL literal has minute resolution
            // and carries slack in the widening direction, so this bound is over-inclusive by
            // design -- the exact comparison happens client-side below, where being wrong is
            // recoverable.
            DateTimeOffset boundary = page!.LastReceived;
            DateTimeOffset effectiveBefore =
                parsedBefore.HasValue && parsedBefore.Value < boundary ? parsedBefore.Value : boundary;

            restrictFilter = MailRestrictFilter.Build(
                unreadOnly,
                fromAddress,
                subjectContains,
                parsedAfter,
                effectiveBefore,
                hasAttachment,
                flaggedOnly,
                pushedDownQuery);
        }

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
                    bool contentIndexAnswered = useContentIndex;
                    string? engineMessage = null;

                    if (restrictFilter != null)
                    {
                        try
                        {
                            restrictedItems = items.Restrict(restrictFilter);
                            scanCount = SafeGetInt(() => restrictedItems.Count);
                            itemsToScan = restrictedItems;
                        }
                        catch (COMException ex)
                        {
                            // Fall back to the unfiltered folder if Restrict is unavailable for
                            // this folder/store type. The client-side checks below are applied
                            // unconditionally, so the result set is identical either way - only
                            // slower, and bounded by ScanSafetyLimit.
                            itemsToScan = items;
                            scanCount = totalItemCount;

                            if (useContentIndex)
                            {
                                // ...except here, where the two paths are *not* equivalent. The
                                // index was asked for and could not answer, so say so: silently
                                // handing back substring semantics under the label the caller asked
                                // for is how a search result becomes quietly wrong.
                                contentIndexAnswered = false;
                                engineMessage =
                                    "This store could not answer from the content index, so the query was run as a "
                                    + "client-side scan instead. That matches substrings rather than whole words and "
                                    + "stops at a scan limit, so a match far back in a very large folder may be missed. "
                                    + $"Outlook reported: {ex.Message}";
                            }
                        }
                    }

                    var result = new MailListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        Query = query,
                        TotalItemCount = totalItemCount,
                        SortedBy = "receivedTime",
                        SortDirection = "descending",
                        SearchEngine = contentIndexAnswered ? "contentIndex" : "clientScan",
                        Message = engineMessage
                    };

                    int scanned = 0;
                    int skipped = 0;

                    // Rolling record of the tied band at the frontier of the scan: the received time
                    // of the last item examined, and every entry id examined at that exact instant.
                    // Received times are not unique, so a cursor pointing only at a timestamp would
                    // either repeat that band or skip the rest of it. Carrying the ids makes the
                    // boundary exact without assuming Outlook orders ties identically twice.
                    DateTimeOffset? boundaryTime = null;
                    var boundaryIds = new List<string>();

                    for (int index = 1;
                         index <= scanCount && scanned < ScanSafetyLimit && result.Messages.Count < boundedMaxCount;
                         index++)
                    {
                        object? rawItem = null;
                        Outlook.MailItem? mail = null;
                        Outlook.MeetingItem? meeting = null;

                        try
                        {
                            rawItem = itemsToScan[index];
                            scanned++;

                            mail = rawItem as Outlook.MailItem;

                            // A meeting invitation is a MeetingItem, not a MailItem. This used to be
                            // `if (mail == null) continue;`, which made every invitation invisible in
                            // every listing with nothing in the response to say so. See #32.
                            meeting = mail == null ? rawItem as Outlook.MeetingItem : null;

                            if (mail == null && meeting == null)
                            {
                                skipped++;
                                continue;
                            }

                            DateTimeOffset? received = mail != null
                                ? SafeGetDateTimeOffset(() => mail.ReceivedTime)
                                : SafeGetDateTimeOffset(() => meeting!.ReceivedTime);
                            string? entryId = mail != null
                                ? SafeGet(() => mail.EntryID)
                                : SafeGet(() => meeting!.EntryID);

                            if (received.HasValue)
                            {
                                DateTimeOffset receivedUtc = received.Value.ToUniversalTime();

                                if (page != null && !page.Includes(receivedUtc, entryId))
                                {
                                    continue;
                                }

                                if (boundaryTime != receivedUtc)
                                {
                                    boundaryTime = receivedUtc;
                                    boundaryIds.Clear();
                                }

                                if (entryId != null)
                                {
                                    boundaryIds.Add(entryId);
                                }
                            }

                            // Applied even when Restrict succeeded. The DASL filter is deliberately
                            // over-inclusive -- it drops any predicate it cannot express exactly
                            // (see MailRestrictFilter) -- so this is what makes the result exact.
                            //
                            // The free-text query is the exception. When the content index answered
                            // it, re-checking it here as a substring would *narrow* the result:
                            // the index legitimately matches a word the substring check would not
                            // see (a different inflection, a term inside an attachment it indexed).
                            // Re-applying it would throw those away and leave the caller believing
                            // the index found nothing.
                            string? clientSideQuery = contentIndexAnswered ? null : query;

                            bool matches = mail != null
                                ? MatchesStructuredFilters(
                                      mail, unreadOnly, fromAddress, subjectContains,
                                      parsedAfter, parsedBefore, hasAttachment, flaggedOnly)
                                  && MatchesQuery(mail, clientSideQuery)
                                : MatchesStructuredFilters(
                                      meeting!, unreadOnly, fromAddress, subjectContains,
                                      parsedAfter, parsedBefore, hasAttachment, flaggedOnly)
                                  && MatchesQuery(meeting!, clientSideQuery);

                            if (!matches)
                            {
                                continue;
                            }

                            result.Messages.Add(mail != null
                                ? CreateMailSummary(mail, includeBodyPreview)
                                : CreateMeetingSummary(meeting!, includeBodyPreview));
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref mail);
                            OutlookInteropRunner.ReleaseComObject(ref meeting);
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ReturnedCount = result.Messages.Count;
                    result.ScannedCount = scanned;
                    result.SkippedItemCount = skipped;
                    // Truncated: there was more to look at than we actually evaluated, whether
                    // because the result cap (maxCount) was hit first or the safety limit was --
                    // either way, "no more results" must not be inferred from this response alone.
                    result.Truncated = scanned < scanCount;

                    // A continuation is only offered when this call actually advanced the frontier.
                    // Handing back a cursor that re-scans the same band would let a caller loop
                    // forever believing it was making progress.
                    if (result.Truncated && boundaryTime.HasValue)
                    {
                        result.NextCursor = MailPageCursor.Encode(fingerprint, boundaryTime.Value, boundaryIds);
                        result.HasMore = true;
                    }

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

    /// <summary>
    /// Reads the caller's <c>bodyFormat</c> argument.
    ///
    /// <para>
    /// Unknown values are refused rather than defaulted. Quietly treating <c>"richtext"</c> as plain
    /// would put the markup the caller expected to be rendered in front of a human as visible tag
    /// soup, and the call would have reported success.
    /// </para>
    /// </summary>
    private static bool TryParseBodyFormat(string? bodyFormat, out bool asHtml, out string? error)
    {
        if (string.IsNullOrWhiteSpace(bodyFormat)
            || bodyFormat.Equals("plain", StringComparison.OrdinalIgnoreCase)
            || bodyFormat.Equals("text", StringComparison.OrdinalIgnoreCase))
        {
            asHtml = false;
            error = null;
            return true;
        }

        if (bodyFormat.Equals("html", StringComparison.OrdinalIgnoreCase))
        {
            asHtml = true;
            error = null;
            return true;
        }

        asHtml = false;
        error = $"bodyFormat '{bodyFormat}' is not supported. Use 'plain' or 'html'.";
        return false;
    }

    /// <summary>
    /// Puts the caller's text above the quoted original in a reply or forward draft, without
    /// destroying the quote.
    ///
    /// <para>
    /// The previous version of this read <c>draftMail.Body</c> - the lossy plain-text projection of a
    /// quoted original that is almost always HTML - and wrote it straight back. That silently
    /// flattened the entire quoted thread: fonts, tables, inline images, links, all of it, replaced by
    /// text. The call reported success and every word was still present, so nothing looked wrong
    /// until someone opened the draft. Writing to <c>HTMLBody</c> instead keeps the original intact.
    /// </para>
    ///
    /// <para>
    /// Plain caller text is HTML-escaped on the way in. Without that, a user writing "profit &lt;
    /// loss" would lose everything after the bracket to an unclosed tag - a silent edit of somebody's
    /// words, which is worse than a visible failure.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void PrependToDraftBody(Outlook.MailItem draftMail, string body, bool asHtml)
    {
        bool draftIsHtml = SafeIsHtmlBody(draftMail);

        if (!draftIsHtml && !asHtml)
        {
            // Nothing to preserve and nothing to render: the plain-text concatenation is correct.
            string existingBody = SafeGet(() => draftMail.Body) ?? string.Empty;
            draftMail.Body = body + Environment.NewLine + Environment.NewLine + existingBody;
            return;
        }

        string addition = asHtml ? body : $"<div>{PlainTextToHtml(body)}</div>";
        string existingHtml = SafeGet(() => draftMail.HTMLBody) ?? string.Empty;

        draftMail.HTMLBody = InsertAtStartOfHtmlBody(existingHtml, addition);
    }

    /// <summary>
    /// Whether the draft is an HTML message. A failure to read the property is treated as "not HTML",
    /// which keeps the caller's text as text rather than risking markup being injected into a message
    /// whose format is unknown.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool SafeIsHtmlBody(Outlook.MailItem mail)
    {
        try
        {
            return mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML;
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>
    /// Escapes plain text for inclusion in an HTML body, keeping the caller's line breaks visible.
    /// </summary>
    private static string PlainTextToHtml(string text)
    {
        string escaped = text
            .Replace("&", "&amp;", StringComparison.Ordinal)
            .Replace("<", "&lt;", StringComparison.Ordinal)
            .Replace(">", "&gt;", StringComparison.Ordinal);

        return escaped
            .Replace("\r\n", "<br>", StringComparison.Ordinal)
            .Replace("\n", "<br>", StringComparison.Ordinal)
            .Replace("\r", "<br>", StringComparison.Ordinal);
    }

    /// <summary>
    /// Inserts content just inside an HTML document's <c>&lt;body&gt;</c> element, so it appears
    /// above everything already there.
    ///
    /// <para>
    /// Concatenating in front of the whole document instead would put the text before
    /// <c>&lt;html&gt;</c>, outside any element. Browsers and Outlook usually recover from that, but
    /// "usually" is not a property worth depending on for the visible content of somebody's mail. If
    /// there is no <c>&lt;body&gt;</c> tag - Outlook does produce bare fragments - prepending is the
    /// correct answer rather than a fallback.
    /// </para>
    /// </summary>
    private static string InsertAtStartOfHtmlBody(string html, string addition)
    {
        if (string.IsNullOrEmpty(html))
        {
            return addition;
        }

        int bodyTag = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);

        if (bodyTag >= 0)
        {
            int tagEnd = html.IndexOf('>', bodyTag);

            if (tagEnd >= 0)
            {
                return html[..(tagEnd + 1)] + addition + html[(tagEnd + 1)..];
            }
        }

        return addition + html;
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
        string bodyFormat,
        Func<Outlook.MailItem, Outlook.MailItem> createDraft)
    {
        if (!TryParseBodyFormat(bodyFormat, out bool asHtml, out string? formatError))
        {
            return new MailDraftResult
            {
                Success = false,
                Saved = false,
                Displayed = false,
                ErrorMessage = formatError
            };
        }

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

                    // Outlook refuses to build a reply or forward from an item that was never sent -
                    // there is nobody to reply to - and reports it as "Could not send the message",
                    // which is nonsense to a caller who is not sending anything and unactionable to
                    // an agent, which will simply retry. Name the actual cause first (#92).
                    if (!SafeGetBool(() => sourceMail.Sent))
                    {
                        return new MailDraftResult
                        {
                            Success = false,
                            Saved = false,
                            Displayed = false,
                            ErrorMessage = "This message is an unsent draft, so Outlook cannot create a reply or forward from it. "
                                + "Reply to or forward a message that was actually sent or received; to edit the draft itself, "
                                + "use mail.set-subject, mail.set-body or mail.set-recipients."
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
                        PrependToDraftBody(draftMail, body, asHtml);
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
        bool? hasAttachment,
        bool flaggedOnly)
    {
        if (unreadOnly && !SafeGetBool(() => mail.UnRead))
        {
            return false;
        }

        // Restrict already excludes these, but the DASL filter is allowed to be over-inclusive by
        // design and a caller can reach this path with no filter pushed down at all. Checking here
        // means "flagged" always means outstanding, never merely "was flagged once".
        if (flaggedOnly && MapFlagStatus(SafeGetInt(() => (int)mail.FlagStatus)) != "flagged")
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

    /// <summary>
    /// Client-side free-text match, used for the fields <c>Restrict</c> cannot filter on.
    /// <para>
    /// The body is matched in full. It used to be truncated to 1200 characters first, which meant a
    /// term further into a long message was silently invisible and the caller was told there was no
    /// such mail (#42). The truncation also bought nothing: the expensive part is the
    /// <c>mail.Body</c> COM call, and by the time the string was cut that had already happened.
    /// </para>
    /// </summary>
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
            || ContainsIgnoreCase(OutlookInteropRunner.NormalizeBodyText(SafeGet(() => mail.Body)), searchText);
    }

    private static bool ContainsIgnoreCase(string? value, string searchText)
        => !string.IsNullOrWhiteSpace(value)
           && value.Contains(searchText, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Resolves the caller's <c>searchMode</c> to a decision about which engine answers the free-text
    /// query (#42).
    /// </summary>
    /// <remarks>
    /// An unrecognised value is an error rather than a default. The two engines do not answer the
    /// same question - the index matches whole words, the scan matches substrings and stops at a
    /// limit - so a typo silently resolving to the default would give the caller results they did not
    /// ask for, in the one place where the difference is invisible: the matches they never see.
    /// </remarks>
    private static bool TryParseSearchMode(string? searchMode, out bool useContentIndex, out string? errorMessage)
    {
        useContentIndex = false;
        errorMessage = null;

        if (string.IsNullOrWhiteSpace(searchMode))
        {
            return true;
        }

        switch (searchMode.Trim().ToLowerInvariant())
        {
            case "clientscan":
            case "client":
                return true;

            case "fulltext":
            case "contentindex":
                useContentIndex = true;
                return true;

            default:
                errorMessage =
                    $"searchMode '{searchMode}' is not recognised. Use 'clientScan' (the default: exact substring "
                    + "matching, bounded by a scan limit) or 'fullText' (Outlook's content index: whole-word matching "
                    + "with no scan limit).";
                return false;
        }
    }

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
            ConversationId = SafeGet(() => mail.ConversationID),
            ConversationTopic = SafeGet(() => mail.ConversationTopic),
            BodyPreview = includeBodyPreview
                ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body))
                : null,
            Categories = ParseCategories(SafeGet(() => mail.Categories)),
            Unread = SafeGetBool(() => mail.UnRead),
            IsDraft = SafeGetBool(() => !mail.Sent && SafeGet(() => mail.MessageClass)?.Contains("IPM.Note", StringComparison.OrdinalIgnoreCase) == true),
            Importance = SafeGetInt(() => (int)mail.Importance),
            AttachmentCount = SafeGetInt(() => mail.Attachments.Count),
            ReceivedTime = SafeGetDateTimeOffset(() => mail.ReceivedTime),
            SentOn = SafeGetDateTimeOffset(() => mail.SentOn),
            FlagStatus = MapFlagStatus(SafeGetInt(() => (int)mail.FlagStatus)),
            FlagRequest = NullIfBlank(SafeGet(() => mail.FlagRequest)),
            FlagDueDate = NormalizeTaskDate(SafeGetDateTimeOffset(() => mail.TaskDueDate))
        };

        summary.AccessDenied = accessDenied.Count > 0 ? accessDenied : null;
        summary.ItemType = "mail";
        return summary;
    }

    /// <summary>
    /// Outlook's stand-in for "this task date was never set". It does not use null, so clearing a
    /// flag means writing this back rather than nulling the field.
    /// </summary>
    private static readonly DateTime NoTaskDate = new(4501, 1, 1);

    /// <summary>
    /// Assigns a task date, tolerating the stores that refuse the write. The flag itself is the
    /// point of the operation; a store that will not record the date should not fail the whole call
    /// and roll nothing back, so the date is best-effort while the status is not.
    /// </summary>
    private static void TrySetTaskDate(Action<DateTime> assign, DateTime value)
    {
        try
        {
            assign(value);
        }
        catch (Exception ex) when (ex is NotImplementedException or COMException)
        {
            // Left unset; the read path reports it as "no due date" rather than inventing one.
        }
    }

    /// <summary>
    /// Maps Outlook's <c>OlFlagStatus</c> onto the wire values. <c>complete</c> is kept distinct from
    /// <c>none</c> deliberately: "I dealt with this" and "this was never flagged" are different
    /// answers to "what still needs attention".
    /// </summary>
    private static string MapFlagStatus(int? flagStatus) => flagStatus switch
    {
        (int)Outlook.OlFlagStatus.olFlagMarked => "flagged",
        (int)Outlook.OlFlagStatus.olFlagComplete => "complete",
        _ => "none"
    };

    private static string? NullIfBlank(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    /// <summary>
    /// Outlook does not leave task dates null when there is no due date - it stores a far-future
    /// sentinel (year 4501). Reporting that verbatim would tell a caller the mail is due in the
    /// 46th century, so it is normalised to "no date".
    /// </summary>
    private static DateTimeOffset? NormalizeTaskDate(DateTimeOffset? value) =>
        value is null || value.Value.Year >= 4000 ? null : value;

    /// <summary>
    /// Builds a listing entry for a meeting request, cancellation or response.
    ///
    /// <para>
    /// A <c>MeetingItem</c> is not a <c>MailItem</c> and shares no interface with it, so this cannot
    /// reuse <see cref="CreateMailSummary"/> - the properties happen to be named the same but are
    /// declared on unrelated COM types. It deliberately reads only the fields that mean the same
    /// thing on both, so a caller can treat a listing uniformly and use <c>itemType</c> to decide
    /// what it may do with an entry. See #32.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static MailSummaryInfo CreateMeetingSummary(Outlook.MeetingItem meeting, bool includeBodyPreview)
    {
        var accessDenied = new List<string>();

        var summary = new MailSummaryInfo
        {
            EntryId = SafeGet(() => meeting.EntryID),
            StoreId = SafeGet(() => meeting.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
            Subject = SafeGet(() => meeting.Subject),
            SenderName = SafeGet(() => meeting.SenderName),
            SenderEmailAddress = SafeGet(() => meeting.SenderEmailAddress, nameof(MailSummaryInfo.SenderEmailAddress), accessDenied),
            // MeetingItem has no To/CC; the attendees live on Recipients. Rendered into To so a
            // listing reads uniformly, rather than leaving a meeting looking like it is addressed to
            // nobody.
            To = SafeGet(() => JoinRecipients(meeting.Recipients), nameof(MailSummaryInfo.To), accessDenied),
            ConversationId = SafeGet(() => meeting.ConversationID),
            ConversationTopic = SafeGet(() => meeting.ConversationTopic),
            BodyPreview = includeBodyPreview
                ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => meeting.Body))
                : null,
            Categories = ParseCategories(SafeGet(() => meeting.Categories)),
            FlagStatus = MapFlagStatus(SafeGetInt(() => (int)meeting.FlagStatus)),
            FlagRequest = NullIfBlank(SafeGet(() => meeting.FlagRequest)),
            Unread = SafeGetBool(() => meeting.UnRead),
            IsDraft = false,
            Importance = SafeGetInt(() => (int)meeting.Importance),
            AttachmentCount = SafeGetInt(() => meeting.Attachments.Count),
            ReceivedTime = SafeGetDateTimeOffset(() => meeting.ReceivedTime),
            SentOn = SafeGetDateTimeOffset(() => meeting.SentOn),
            ItemType = ClassifyMeetingItem(SafeGet(() => meeting.MessageClass))
        };

        summary.AccessDenied = accessDenied.Count > 0 ? accessDenied : null;
        return summary;
    }

    /// <summary>
    /// Maps a scheduling message class onto the distinction that actually changes what a caller may
    /// do: an invitation can be responded to, a cancellation cannot, and a response is somebody
    /// else's answer to an invitation you sent.
    /// </summary>
    private static string ClassifyMeetingItem(string? messageClass)
    {
        if (string.IsNullOrWhiteSpace(messageClass))
        {
            return "other";
        }

        if (messageClass.StartsWith("IPM.Schedule.Meeting.Canceled", StringComparison.OrdinalIgnoreCase))
        {
            return "meetingCancellation";
        }

        if (messageClass.StartsWith("IPM.Schedule.Meeting.Resp", StringComparison.OrdinalIgnoreCase))
        {
            return "meetingResponse";
        }

        return messageClass.StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase)
            ? "meetingRequest"
            : "other";
    }

    /// <summary>
    /// The structured-filter check for meeting items. Mirrors the <c>MailItem</c> overload exactly; the duplication exists because <c>MeetingItem</c> and <c>MailItem</c> are unrelated
    /// COM types with coincidentally identical property names.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool MatchesStructuredFilters(
        Outlook.MeetingItem meeting,
        bool unreadOnly,
        string? fromAddress,
        string? subjectContains,
        DateTimeOffset? receivedAfter,
        DateTimeOffset? receivedBefore,
        bool? hasAttachment,
        bool flaggedOnly)
    {
        if (unreadOnly && !SafeGetBool(() => meeting.UnRead))
        {
            return false;
        }

        // A meeting request carries the same follow-up flag as a message, so it is filtered the same
        // way rather than being dropped from flagged results for being a different item type.
        if (flaggedOnly && MapFlagStatus(SafeGetInt(() => (int)meeting.FlagStatus)) != "flagged")
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(subjectContains)
            && !ContainsIgnoreCase(SafeGet(() => meeting.Subject), subjectContains.Trim()))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(fromAddress))
        {
            string needle = fromAddress.Trim();
            if (!ContainsIgnoreCase(SafeGet(() => meeting.SenderEmailAddress), needle)
                && !ContainsIgnoreCase(SafeGet(() => meeting.SenderName), needle))
            {
                return false;
            }
        }

        if (hasAttachment.HasValue && SafeGetInt(() => meeting.Attachments.Count) > 0 != hasAttachment.Value)
        {
            return false;
        }

        if (receivedAfter.HasValue || receivedBefore.HasValue)
        {
            DateTimeOffset? received = SafeGetDateTimeOffset(() => meeting.ReceivedTime);
            if (received == null)
            {
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

    /// <summary>Free-text match for meeting items. See <see cref="MatchesQuery(Outlook.MailItem, string?)"/>.</summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool MatchesQuery(Outlook.MeetingItem meeting, string? query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return true;
        }

        string searchText = query.Trim();
        return ContainsIgnoreCase(SafeGet(() => meeting.Subject), searchText)
            || ContainsIgnoreCase(SafeGet(() => meeting.SenderName), searchText)
            || ContainsIgnoreCase(SafeGet(() => meeting.SenderEmailAddress), searchText)
            || ContainsIgnoreCase(SafeGet(() => JoinRecipients(meeting.Recipients)), searchText)
            || ContainsIgnoreCase(OutlookInteropRunner.NormalizeBodyText(SafeGet(() => meeting.Body)), searchText);
    }

    /// <summary>
    /// Renders a meeting's attendees into a single display string, so a meeting entry in a listing
    /// carries the same "who is this addressed to" information a message does.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? JoinRecipients(Outlook.Recipients? recipients)
    {
        if (recipients == null)
        {
            return null;
        }

        var names = new List<string>();

        try
        {
            for (int index = 1; index <= recipients.Count; index++)
            {
                Outlook.Recipient? recipient = null;

                try
                {
                    recipient = recipients[index];
                    string? name = SafeGet(() => recipient.Name);

                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref recipient);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref recipients);
        }

        return names.Count > 0 ? string.Join("; ", names) : null;
    }
}




