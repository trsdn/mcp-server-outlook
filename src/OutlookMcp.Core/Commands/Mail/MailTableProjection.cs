using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// Builds listing rows from an Outlook <c>Table</c> rowset instead of opening a <c>MailItem</c> per
/// result (#27).
///
/// <para>
/// <b>Why this exists.</b> A listing used to hydrate every message it returned and read fifteen
/// properties off it, each one a cross-process COM call into Outlook. Measured on this project's
/// reference mailbox, twenty-five rows cost about 9.2 seconds that way and about 0.43 seconds
/// through a <c>Table</c> - a rowset that carries only the columns asked for and never materialises
/// an item. The work is the same; the number of round trips is not.
/// </para>
///
/// <para>
/// <b>What a table cannot do, and why that is stated rather than papered over.</b> Two fields are
/// genuinely unavailable from a rowset: the message body (the <c>Body</c> column can be added but
/// throws when read) and the exact attachment count (only a has-attachments bit exists). So this
/// projection is used only when the request needs neither, and <see cref="MailSummaryInfo.HasAttachment"/>
/// carries the bit the rowset does have while <see cref="MailSummaryInfo.AttachmentCount"/> is left
/// absent. Reporting <c>0</c> attachments for a message with three, under <c>success: true</c>, is
/// precisely the confidently-wrong answer this repository keeps having to remove.
/// </para>
///
/// <para>
/// <b>Column naming is not a free choice.</b> Outlook accepts both its own property names and DASL
/// tags here, and they do not behave the same. Verified against a live mailbox:
/// </para>
/// <list type="bullet">
/// <item><description>
/// DASL date tags such as <c>urn:schemas:httpmail:datereceived</c> return <b>UTC</b>, while the
/// Outlook name <c>ReceivedTime</c> returns the same local wall-clock value as
/// <c>MailItem.ReceivedTime</c>, to the tick. Mixing the two would shift every timestamp by the
/// machine's UTC offset - and would silently corrupt the paging cursor, which is a keyset walk over
/// exactly this value.
/// </description></item>
/// <item><description>
/// <c>urn:schemas:office:office#Keywords</c> cannot be added as a column at all ("does not support
/// this operation"), whereas the Outlook name <c>Categories</c> returns the same comma-joined string
/// <c>MailItem.Categories</c> does.
/// </description></item>
/// <item><description>
/// <c>ConversationID</c> is the one field with no usable name: it must be fetched as the binary
/// proptag <c>0x30130102</c> and hex-encoded, which reproduces <c>MailItem.ConversationID</c> byte
/// for byte.
/// </description></item>
/// <item><description>
/// The <c>EntryID</c> column is <b>not</b> the entry id the rest of this surface uses. It returns the
/// provider's <i>short-term</i> id (<c>PR_ENTRYID</c>), while <c>MailItem.EntryID</c> returns the
/// long-term one; on a cached Exchange mailbox they are 24 and 70 bytes and share only a prefix.
/// Both resolve through <c>GetItemFromID</c>, which is exactly what makes the mistake dangerous - a
/// listing would look entirely correct while handing back ids that never compare equal to the ones
/// <c>mail.read</c> and <c>get-conversation</c> report for the same message. The long-term id is
/// taken from <c>PR_LONGTERM_ENTRYID_FROM_TABLE</c> (<c>0x66700102</c>), which MAPI publishes for
/// this purpose and which was verified byte-for-byte equal to <c>MailItem.EntryID</c>.
/// </description></item>
/// <item><description>
/// A column that can be <i>added</i> is not thereby a column that returns data - <c>Columns.Add</c>
/// succeeds for properties that then read as null or throw. Every column below was confirmed against
/// an item that actually carried the value, not merely by adding it.
/// </description></item>
/// </list>
///
/// <para>
/// <b>Sorting.</b> <c>Table.Sort</c> refuses the DASL date tags outright, so the ordering the cursor
/// depends on is established with the Outlook name <c>[ReceivedTime]</c>. A table that could not be
/// sorted is not used at all: unordered rows would produce a keyset cursor that skips mail without
/// anything in the response saying so.
/// </para>
/// </summary>
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
internal static class MailTableProjection
{
    /// <summary>Value of <see cref="MailListResult.Projection"/> for rows built here.</summary>
    public const string ProjectionName = "table";

    /// <summary>Value of <see cref="MailListResult.Projection"/> for rows built by opening items.</summary>
    public const string ItemProjectionName = "item";

    private const string EntryIdColumn = "http://schemas.microsoft.com/mapi/proptag/0x66700102";
    private const string SubjectColumn = "Subject";
    private const string SenderNameColumn = "SenderName";
    private const string SenderEmailColumn = "SenderEmailAddress";
    private const string ToColumn = "To";
    private const string CcColumn = "CC";
    private const string ReceivedTimeColumn = "ReceivedTime";
    private const string SentOnColumn = "SentOn";
    private const string ImportanceColumn = "Importance";
    private const string UnreadColumn = "UnRead";
    private const string CategoriesColumn = "Categories";
    private const string ConversationTopicColumn = "ConversationTopic";
    private const string MessageClassColumn = "MessageClass";
    private const string FlagStatusColumn = "FlagStatus";
    private const string FlagRequestColumn = "FlagRequest";
    private const string TaskDueDateColumn = "TaskDueDate";

    /// <summary><c>PR_CONVERSATION_ID</c>, binary. There is no Outlook property name for it.</summary>
    private const string ConversationIdColumn = "http://schemas.microsoft.com/mapi/proptag/0x30130102";

    /// <summary>
    /// <c>PR_MESSAGE_FLAGS</c>. Carries the three bits a rowset has no friendlier source for:
    /// read state, unsent (draft) state, and has-attachments.
    /// </summary>
    private const string MessageFlagsColumn = "http://schemas.microsoft.com/mapi/proptag/0x0E070003";

    private const int MsgFlagUnsent = 0x8;
    private const int MsgFlagHasAttach = 0x10;

    /// <summary>The ordering every listing is sorted by, and the one the paging cursor walks.</summary>
    private const string SortProperty = "[ReceivedTime]";

    private static readonly string[] Columns =
    [
        EntryIdColumn,
        SubjectColumn,
        SenderNameColumn,
        SenderEmailColumn,
        ToColumn,
        CcColumn,
        ReceivedTimeColumn,
        SentOnColumn,
        ImportanceColumn,
        UnreadColumn,
        CategoriesColumn,
        ConversationTopicColumn,
        MessageClassColumn,
        FlagStatusColumn,
        FlagRequestColumn,
        TaskDueDateColumn,
        ConversationIdColumn,
        MessageFlagsColumn
    ];

    /// <summary>
    /// Restricts <paramref name="table"/> to the projected columns and sorts it newest-first.
    ///
    /// <para>
    /// Throws rather than degrading if either step fails. The caller treats that as "this store
    /// cannot answer from a rowset" and falls back to opening items, which is slower but complete;
    /// carrying on with a partially configured table would hand back a listing that is missing
    /// columns or is not in the order its own cursor claims.
    /// </para>
    /// </summary>
    public static void Configure(Outlook.Table table)
    {
        Outlook.Columns? columns = null;

        try
        {
            // Hoisted into a local rather than dereferenced per column. `table.Columns` materialises
            // a fresh RCW on every access, so the obvious loop would create nineteen of them and
            // release none - on the hottest path in this file, inside a long-lived server process.
            columns = table.Columns;
            columns.RemoveAll();

            foreach (string column in Columns)
            {
                columns.Add(column);
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref columns);
        }

        table.Sort(SortProperty, Outlook.OlSortOrder.olDescending);
    }

    /// <summary>Received time of the row, as the cursor and the client-side date filters see it.</summary>
    public static DateTimeOffset? ReadReceived(Outlook.Row row)
        => ReadDateTimeOffset(row, ReceivedTimeColumn);

    /// <summary>
    /// The long-term entry id, the one every other action on this surface reports and accepts, or
    /// <see langword="null"/> when this store does not publish it.
    ///
    /// <para>
    /// A null here is treated by the caller as "this store cannot be listed from a rowset" rather
    /// than being quietly replaced with the short-term <c>EntryID</c> column. The two resolve
    /// equally well but never compare equal, so substituting one would make the same message look
    /// like two different messages depending on which action reported it.
    /// </para>
    /// </summary>
    public static string? ReadEntryId(Outlook.Row row) => ReadHex(row, EntryIdColumn);

    /// <summary>
    /// Builds one listing row from a row already known to be a modelled item type.
    /// </summary>
    /// <param name="row">The rowset row.</param>
    /// <param name="storeId">Store the folder being listed belongs to; every row in it shares one.</param>
    /// <param name="itemType">The value <see cref="ClassifyRow"/> returned for this row.</param>
    /// <param name="entryId">The long-term entry id <see cref="ReadEntryId"/> returned for this row.</param>
    public static MailSummaryInfo CreateSummary(Outlook.Row row, string? storeId, string itemType, string entryId)
    {
        int messageFlags = ReadInt(row, MessageFlagsColumn) ?? 0;
        bool isMail = string.Equals(itemType, "mail", StringComparison.Ordinal);

        return new MailSummaryInfo
        {
            EntryId = entryId,
            StoreId = storeId,
            Subject = ReadString(row, SubjectColumn),
            SenderName = ReadString(row, SenderNameColumn),
            SenderEmailAddress = ReadString(row, SenderEmailColumn),
            To = ReadString(row, ToColumn),
            Cc = ReadString(row, CcColumn),
            ConversationId = ReadHex(row, ConversationIdColumn),
            ConversationTopic = ReadString(row, ConversationTopicColumn),
            Categories = ParseCategories(ReadString(row, CategoriesColumn)),
            // PR_MESSAGE_FLAGS rather than the UnRead column, so read state, draft state and
            // attachments all come from one consistent source. The two agree; using one avoids a
            // listing that could contradict itself.
            Unread = (messageFlags & 1) == 0,
            IsDraft = isMail && (messageFlags & MsgFlagUnsent) != 0,
            HasAttachment = (messageFlags & MsgFlagHasAttach) != 0,
            Importance = ReadInt(row, ImportanceColumn) ?? 0,
            ReceivedTime = ReadDateTimeOffset(row, ReceivedTimeColumn),
            SentOn = ReadDateTimeOffset(row, SentOnColumn),
            FlagStatus = MapFlagStatus(ReadInt(row, FlagStatusColumn)),
            FlagRequest = NullIfBlank(ReadString(row, FlagRequestColumn)),
            FlagDueDate = NormalizeTaskDate(ReadDateTimeOffset(row, TaskDueDateColumn)),
            ItemType = itemType
        };
    }

    /// <summary>
    /// The client-side narrowing applied on top of the pushed-down DASL filter, mirroring the
    /// item-based overload field for field.
    ///
    /// <para>
    /// It is applied even when <c>Restrict</c> succeeded, for the same reason it is there: the DASL
    /// filter is deliberately over-inclusive - it drops any predicate it cannot express exactly -
    /// so this is what makes the result set exact.
    /// </para>
    /// </summary>
    public static bool MatchesStructuredFilters(
        Outlook.Row row,
        bool unreadOnly,
        string? fromAddress,
        string? subjectContains,
        DateTimeOffset? receivedAfter,
        DateTimeOffset? receivedBefore,
        bool? hasAttachment,
        bool flaggedOnly)
    {
        int messageFlags = ReadInt(row, MessageFlagsColumn) ?? 0;

        if (unreadOnly && (messageFlags & 1) != 0)
        {
            return false;
        }

        if (flaggedOnly && MapFlagStatus(ReadInt(row, FlagStatusColumn)) != "flagged")
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(subjectContains)
            && !ContainsIgnoreCase(ReadString(row, SubjectColumn), subjectContains.Trim()))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(fromAddress))
        {
            string needle = fromAddress.Trim();
            if (!ContainsIgnoreCase(ReadString(row, SenderEmailColumn), needle)
                && !ContainsIgnoreCase(ReadString(row, SenderNameColumn), needle))
            {
                return false;
            }
        }

        if (hasAttachment.HasValue && ((messageFlags & MsgFlagHasAttach) != 0) != hasAttachment.Value)
        {
            return false;
        }

        if (receivedAfter.HasValue || receivedBefore.HasValue)
        {
            DateTimeOffset? received = ReadDateTimeOffset(row, ReceivedTimeColumn);
            if (received == null)
            {
                // Mirrors the item path: a message whose received time cannot be read cannot be
                // shown to satisfy a date window, so it is excluded rather than guessed at.
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
    /// Decides what a folder entry is, returning <see langword="null"/> for anything this surface
    /// does not model - an appointment or a delivery report sitting in a mail folder, say. The caller
    /// counts those into <c>skippedItemCount</c> exactly as the item-based path does, rather than
    /// dropping them silently.
    /// </summary>
    public static string? ClassifyRow(Outlook.Row row)
        => ClassifyMessageClass(ReadString(row, MessageClassColumn));

    /// <summary>
    /// Decides what a folder entry is from its message class, returning <see langword="null"/> for
    /// anything this surface does not model.
    ///
    /// <para>
    /// This reproduces the item-based path's type tests rather than inventing a new taxonomy: there,
    /// a row became a result only if it cast to <c>MailItem</c> (message classes under
    /// <c>IPM.Note</c>) or <c>MeetingItem</c> (<c>IPM.Schedule.Meeting</c>), and everything else -
    /// appointments, posts, delivery reports - fell through to the skipped count.
    /// </para>
    /// </summary>
    public static string? ClassifyMessageClass(string? messageClass)
    {
        if (string.IsNullOrWhiteSpace(messageClass))
        {
            return null;
        }

        if (messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
        {
            return "mail";
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
            : null;
    }

    /// <summary>
    /// Reads one column, treating a failure as an absent value.
    ///
    /// <para>
    /// A <c>Row</c> throws rather than returning null for some properties on some item types, and one
    /// unreadable column on one row must not fail a whole listing. This is the same narrow
    /// optional-property-read tolerance the item path uses, not exception suppression around real
    /// work: the row itself has already been fetched by the time this runs.
    /// </para>
    /// </summary>
    private static object? Read(Outlook.Row row, string column)
    {
        try
        {
            return row[column];
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static string? ReadString(Outlook.Row row, string column)
        => Read(row, column) as string;

    private static int? ReadInt(Outlook.Row row, string column)
        => Read(row, column) is int value ? value : null;

    /// <summary>
    /// Reads a date column. <c>DateTime.Kind</c> comes back as <c>Unspecified</c> carrying local wall
    /// clock, matching what the item path gets from <c>MailItem.ReceivedTime</c>, so it is converted
    /// the same way.
    /// </summary>
    private static DateTimeOffset? ReadDateTimeOffset(Outlook.Row row, string column)
    {
        if (Read(row, column) is not DateTime value || value == default)
        {
            return null;
        }

        return new DateTimeOffset(value);
    }

    /// <summary>
    /// Renders a binary column as the uppercase hex string Outlook's own object model reports for
    /// the same property. Verified equal to <c>MailItem.ConversationID</c> on a live mailbox; a
    /// lowercase or delimited rendering would look plausible and never match a conversation.
    /// </summary>
    private static string? ReadHex(Outlook.Row row, string column)
        => Read(row, column) is byte[] { Length: > 0 } bytes ? Convert.ToHexString(bytes) : null;

    private static List<string> ParseCategories(string? categories)
        => string.IsNullOrWhiteSpace(categories)
            ? []
            : [.. categories.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)];

    private static string MapFlagStatus(int? flagStatus) => flagStatus switch
    {
        (int)Outlook.OlFlagStatus.olFlagMarked => "flagged",
        (int)Outlook.OlFlagStatus.olFlagComplete => "complete",
        _ => "none"
    };

    private static string? NullIfBlank(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value;

    /// <summary>
    /// Outlook stores a far-future sentinel rather than null for "no due date"; reporting it
    /// verbatim would tell a caller the mail is due in the 46th century.
    /// </summary>
    private static DateTimeOffset? NormalizeTaskDate(DateTimeOffset? value)
        => value is null || value.Value.Year >= 4000 ? null : value;

    private static bool ContainsIgnoreCase(string? value, string searchText)
        => !string.IsNullOrWhiteSpace(value)
           && value.Contains(searchText, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Formats the scan-limit explanation shared by both projections, so the two paths cannot drift
    /// into describing the same truncation differently.
    /// </summary>
    public static string DescribeFallback(string reason)
        => string.Format(
            CultureInfo.InvariantCulture,
            "This folder could not be listed from an Outlook table rowset, so each message was opened "
            + "instead. The results are the same; the call was slower. Outlook reported: {0}",
            reason);
}
