using System.Text.Json.Serialization;

namespace OutlookMcp.Core.Models;

/// <summary>
/// Base result type for all Core operations.
/// Exceptions propagate naturally — batch.Execute() re-throws them via TaskCompletionSource.
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

/// <summary>
/// Result for rename operations
/// </summary>
public class RenameResult : ResultBase
{
    public string ObjectType { get; set; } = string.Empty;
    public string OldName { get; set; } = string.Empty;
    public string NewName { get; set; } = string.Empty;
}

// ── File / Session ────────────────────────────────────────

public class FileValidationInfo : ResultBase
{
    public bool Exists { get; set; }
    public string FileName { get; set; } = string.Empty;
    public long FileSizeBytes { get; set; }
    public bool IsReadOnly { get; set; }
    public bool IsMacroEnabled { get; set; }
    public int SlideCount { get; set; }
}

// ── Slide ─────────────────────────────────────────────────

public class SlideListResult : ResultBase
{
    public List<SlideInfo> Slides { get; set; } = [];
}

public class SlideInfo
{
    public int SlideIndex { get; set; }
    public int SlideNumber { get; set; }
    public string SlideId { get; set; } = string.Empty;
    public string LayoutName { get; set; } = string.Empty;
    public string MasterName { get; set; } = string.Empty;
    public int ShapeCount { get; set; }
    public bool HasNotes { get; set; }
    public bool HasAnimations { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }
}

public class SlideDetailResult : ResultBase
{
    public SlideInfo? Slide { get; set; }
    public List<ShapeInfo> Shapes { get; set; } = [];
}

// ── Shape ─────────────────────────────────────────────────

public class ShapeListResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<ShapeInfo> Shapes { get; set; } = [];
}

public class ShapeInfo
{
    public int ShapeId { get; set; }
    public string Name { get; set; } = string.Empty;
    public string ShapeType { get; set; } = string.Empty;
    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
    public int ZOrderPosition { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AlternativeText { get; set; }

    public bool HasTextFrame { get; set; }
    public bool HasTable { get; set; }
    public bool HasChart { get; set; }
    public bool IsGroup { get; set; }
    public bool IsPlaceholder { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PlaceholderType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<ShapeInfo>? GroupItems { get; set; }
}

public class ShapeDetailResult : ResultBase
{
    public ShapeInfo? Shape { get; set; }
}

// ── Text ──────────────────────────────────────────────────

public class TextResult : ResultBase
{
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public List<TextParagraphInfo> Paragraphs { get; set; } = [];
}

public class TextParagraphInfo
{
    public int Index { get; set; }
    public string Text { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Alignment { get; set; }

    public List<TextRunInfo> Runs { get; set; } = [];
}

public class TextRunInfo
{
    public string Text { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? FontSize { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Bold { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Italic { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; set; }
}

// ── Table (in shapes) ────────────────────────────────────

public class SlideTableResult : ResultBase
{
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public List<List<string?>> Data { get; set; } = [];
}

// ── Master / Layout ───────────────────────────────────────

public class MasterListResult : ResultBase
{
    public List<MasterInfo> Masters { get; set; } = [];
}

public class MasterInfo
{
    public string Name { get; set; } = string.Empty;
    public List<LayoutInfo> Layouts { get; set; } = [];
}

public class LayoutInfo
{
    public string Name { get; set; } = string.Empty;
    public int Index { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MatchingName { get; set; }
}

// ── Notes ─────────────────────────────────────────────────

public class NotesResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string Text { get; set; } = string.Empty;
}

// ── Transition ────────────────────────────────────────────

public class TransitionResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string TransitionType { get; set; } = string.Empty;
    public float Duration { get; set; }
    public bool AdvanceOnClick { get; set; }
    public float AdvanceAfterTime { get; set; }
}

// ── Animation ─────────────────────────────────────────────

public class AnimationListResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<AnimationInfo> Animations { get; set; } = [];
}

public class AnimationInfo
{
    public int Index { get; set; }
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string EffectType { get; set; } = string.Empty;
    public string Timing { get; set; } = string.Empty;
    public float Duration { get; set; }
    public float Delay { get; set; }
}

// ── Export ─────────────────────────────────────────────────

public class ExportResult : ResultBase
{
    public string OutputPath { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
}

// ── Outlook application / folder / mail ─────────────────────

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

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ItemCount { get; set; }
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

// ── Chart ──────────────────────────────────────────────────

public class ChartInfoResult : ResultBase
{
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public int ChartType { get; set; }
    public string ChartTypeName { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; set; }

    public bool HasLegend { get; set; }
    public int SeriesCount { get; set; }
}

// ── Design / Theme ────────────────────────────────────────

public class DesignListResult : ResultBase
{
    public List<DesignInfo> Designs { get; set; } = [];
}

public class DesignInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int LayoutCount { get; set; }
}

public class ThemeColorResult : ResultBase
{
    public string DesignName { get; set; } = string.Empty;
    public Dictionary<string, string> Colors { get; set; } = [];
}

// ── Theme Fonts ──────────────────────────────────────────

public class ThemeFontResult : ResultBase
{
    public string DesignName { get; set; } = string.Empty;
    public string HeadingFont { get; set; } = string.Empty;
    public string BodyFont { get; set; } = string.Empty;
}

// ── Slideshow ─────────────────────────────────────────────

public class SlideshowInfoResult : ResultBase
{
    public bool IsRunning { get; set; }
    public int CurrentSlide { get; set; }
    public int TotalSlides { get; set; }
}

// ── VBA ───────────────────────────────────────────────────

public class VbaModuleListResult : ResultBase
{
    public List<VbaModuleInfo> Modules { get; set; } = [];
}

public class VbaModuleInfo
{
    public string Name { get; set; } = string.Empty;
    public int ModuleType { get; set; }
    public string ModuleTypeName { get; set; } = string.Empty;
    public int LineCount { get; set; }
}

public class VbaModuleCodeResult : ResultBase
{
    public string ModuleName { get; set; } = string.Empty;
    public string Code { get; set; } = string.Empty;
    public int LineCount { get; set; }
}

// ── Window ────────────────────────────────────────────────

public class WindowInfoResult : ResultBase
{
    public int WindowState { get; set; }
    public string WindowStateName { get; set; } = string.Empty;
    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
}

// ── Hyperlink ─────────────────────────────────────────────

public class HyperlinkResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public bool HasHyperlink { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ScreenTip { get; set; }
}

public class HyperlinkListResult : ResultBase
{
    public List<HyperlinkInfo> Hyperlinks { get; set; } = [];
}

public class HyperlinkInfo
{
    public int Index { get; set; }
    public string Address { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ScreenTip { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SlideIndex { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeName { get; set; }
}

// ── Section ───────────────────────────────────────────────

public class SectionListResult : ResultBase
{
    public List<SectionInfo> Sections { get; set; } = [];
}

public class SectionInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int FirstSlideIndex { get; set; }
    public int SlideCount { get; set; }
}

// ── Document Properties ───────────────────────────────────

public class DocumentPropertyResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Author { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Keywords { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Comments { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Company { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Category { get; set; }
}

// ── Media ─────────────────────────────────────────────────

public class MediaInfoResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string MediaType { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceFile { get; set; }

    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
}

// ── Comment ──────────────────────────────────────────────

public class CommentListResult : ResultBase
{
    public List<CommentInfo> Comments { get; set; } = [];
}

public class CommentInfo
{
    public int SlideIndex { get; set; }
    public int CommentIndex { get; set; }
    public string Author { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public float Left { get; set; }
    public float Top { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DateTime { get; set; }
}

// ── Placeholder ──────────────────────────────────────────

public class PlaceholderListResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<PlaceholderInfo> Placeholders { get; set; } = [];
}

public class PlaceholderInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int PlaceholderType { get; set; }
    public string PlaceholderTypeName { get; set; } = string.Empty;
    public bool HasTextFrame { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; set; }
}

// ── Background ───────────────────────────────────────────

public class BackgroundResult : ResultBase
{
    public int SlideIndex { get; set; }
    public bool FollowMasterBackground { get; set; }
    public string FillType { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; set; }
}

// ── Header/Footer ────────────────────────────────────────

public class HeaderFooterResult : ResultBase
{
    public bool ShowFooter { get; set; }
    public bool ShowSlideNumber { get; set; }
    public bool ShowDate { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FooterText { get; set; }
}

// ── SmartArt ─────────────────────────────────────────────

public class SmartArtInfoResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string LayoutName { get; set; } = string.Empty;
    public List<SmartArtNodeInfo> Nodes { get; set; } = [];
}

public class SmartArtNodeInfo
{
    public int Index { get; set; }
    public string Text { get; set; } = string.Empty;
    public int Level { get; set; }
}

// ── Custom Show ──────────────────────────────────────────

public class CustomShowListResult : ResultBase
{
    public List<CustomShowInfo> Shows { get; set; } = [];
}

public class CustomShowInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int SlideCount { get; set; }
    public List<int> SlideIds { get; set; } = [];
}

// ── Page Setup ───────────────────────────────────────────

public class PageSetupResult : ResultBase
{
    public float SlideWidth { get; set; }
    public float SlideHeight { get; set; }
    public int SlideOrientation { get; set; }
    public int NotesOrientation { get; set; }
}

// ── Tags ─────────────────────────────────────────────────

public class TagListResult : ResultBase
{
    public int SlideIndex { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeName { get; set; }

    public List<TagInfo> Tags { get; set; } = [];
}

public class TagInfo
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

// ── Color Scheme ─────────────────────────────────────────

public class ColorSchemeListResult : ResultBase
{
    public List<ColorSchemeInfo> ColorSchemes { get; set; } = [];
}

public class ColorSchemeInfo
{
    public int Index { get; set; }
    public Dictionary<string, string> Colors { get; set; } = [];
}

// ── Accessibility ────────────────────────────────────────

public class AccessibilityAuditResult : OperationResult
{
    public int TotalSlides { get; set; }
    public int IssueCount { get; set; }
    public List<AccessibilityIssue> Issues { get; set; } = [];
}

public class AccessibilityIssue
{
    public int SlideIndex { get; set; }
    public string IssueType { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeName { get; set; }

    public string Description { get; set; } = string.Empty;
}

public class ReadingOrderResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<ReadingOrderEntry> Shapes { get; set; } = [];
}

public class ReadingOrderEntry
{
    public int Position { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string ShapeType { get; set; } = string.Empty;
    public int ZOrderPosition { get; set; }
}

// ── Design Catalog ───────────────────────────────────────

public class ArchetypeListResult : ResultBase
{
    public List<ArchetypeListItem> Archetypes { get; set; } = [];
}

public class ArchetypeListItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string When { get; set; } = string.Empty;
    public List<string> BestDensity { get; set; } = [];
    public List<string> Variants { get; set; } = [];
    public string ExampleTitle { get; set; } = string.Empty;
    public bool HasCuratedLayoutGuidance { get; set; }
    public int ObservedSlideCount { get; set; }
    public int ObservedSubtypeCount { get; set; }
    public List<string> ObservedExampleSlides { get; set; } = [];
}

public class ArchetypeDetailResult : ResultBase
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string When { get; set; } = string.Empty;
    public List<string> BestDensity { get; set; } = [];
    public List<string> Variants { get; set; } = [];
    public bool HasCuratedLayoutGuidance { get; set; }
    public int ObservedSlideCount { get; set; }
    public List<string> ObservedExampleSlides { get; set; } = [];
    public List<ReferenceSlideInfo> ObservedExamples { get; set; } = [];
    public List<ReferenceSubtypeInfo> ObservedSubtypes { get; set; } = [];
    public List<ReferenceMisbucketedSampleInfo> AuditSamples { get; set; } = [];
    public string Detail { get; set; } = string.Empty;
}

public class ReferenceSubtypeInfo
{
    public string SubArchetypeId { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public List<string> HeuristicPhrases { get; set; } = [];
    public int Count { get; set; }
    public List<string> ExampleSlides { get; set; } = [];
    public List<ReferenceSlideInfo> ExampleDetails { get; set; } = [];
}

public class ReferenceMisbucketedSampleInfo
{
    public string ReferenceId { get; set; } = string.Empty;
    public string CurrentArchetypeId { get; set; } = string.Empty;
    public string SuggestedArchetypeId { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
}

public class ReferenceSlideInfo
{
    public string Id { get; set; } = string.Empty;
    public string ArchetypeId { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubArchetypeId { get; set; }

    public string Rationale { get; set; } = string.Empty;
}

public class PaletteListResult : ResultBase
{
    public List<PaletteListItem> Palettes { get; set; } = [];
}

public class PaletteListItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string BestFor { get; set; } = string.Empty;
}

public class PaletteDetailResult : ResultBase
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string BestFor { get; set; } = string.Empty;
    public Dictionary<string, string> Colors { get; set; } = [];
}

public class StyleProfileListResult : ResultBase
{
    public List<StyleProfileListItem> Profiles { get; set; } = [];
}

public class StyleProfileListItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string BestFor { get; set; } = string.Empty;
    public string ColorScheme { get; set; } = string.Empty;
}

public class StyleProfileDetailResult : ResultBase
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string BestFor { get; set; } = string.Empty;
    public string ColorScheme { get; set; } = string.Empty;
    public string Font { get; set; } = string.Empty;
    public string TitleStyle { get; set; } = string.Empty;
    public int TitleSize { get; set; }
    public int BodySize { get; set; }
    public int FootnoteSize { get; set; }
    public string BulletsPerSlide { get; set; } = string.Empty;
    public string WordsPerBullet { get; set; } = string.Empty;
    public string ContentDensity { get; set; } = string.Empty;
    public List<string> PreferredArchetypes { get; set; } = [];
    public string Whitespace { get; set; } = string.Empty;
    public string Background { get; set; } = string.Empty;
    public string ChartStyle { get; set; } = string.Empty;
    public string SpecialRules { get; set; } = string.Empty;
}

public class LayoutGridResult : ResultBase
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string BestFor { get; set; } = string.Empty;
    public List<LayoutZone> Zones { get; set; } = [];
}

public class LayoutZone
{
    public string Name { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? X { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Y { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? W { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? H { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Description { get; set; }
}

public class LayoutGridListResult : ResultBase
{
    public List<LayoutGridListItem> Grids { get; set; } = [];
}

public class LayoutGridListItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string BestFor { get; set; } = string.Empty;
}

public class DensityProfileResult : ResultBase
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string UsedFor { get; set; } = string.Empty;
    public string Audience { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string TextVolume { get; set; } = string.Empty;
    public string ElementCount { get; set; } = string.Empty;
    public string DataGranularity { get; set; } = string.Empty;
    public string AnnotationDepth { get; set; } = string.Empty;
    public string SourceCompleteness { get; set; } = string.Empty;
    public string WhiteSpaceRatio { get; set; } = string.Empty;
    public string Character { get; set; } = string.Empty;
    public List<string> BestArchetypes { get; set; } = [];
}

public class DensityProfileListResult : ResultBase
{
    public List<DensityProfileListItem> Profiles { get; set; } = [];
}

public class DensityProfileListItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string UsedFor { get; set; } = string.Empty;
}

public class ContextModelResult : ResultBase
{
    public List<MeetingTypeInfo> MeetingTypes { get; set; } = [];
    public List<AudienceLevelInfo> AudienceLevels { get; set; } = [];
    public List<ConsumptionModeInfo> ConsumptionModes { get; set; } = [];
    public string DefaultDensity { get; set; } = string.Empty;
}

public class MeetingTypeInfo
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Audience { get; set; } = string.Empty;
    public string TimePerSlide { get; set; } = string.Empty;
    public string Goal { get; set; } = string.Empty;
    public string DecisionPressure { get; set; } = string.Empty;
    public string PrimaryMode { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SecondaryMode { get; set; }

    public string DefaultDensity { get; set; } = string.Empty;
}

public class AudienceLevelInfo
{
    public string Id { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public string Roles { get; set; } = string.Empty;
    public string PreferredDensity { get; set; } = string.Empty;
    public string WantsToSee { get; set; } = string.Empty;
}

public class ConsumptionModeInfo
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public bool SpeakerPresent { get; set; }
    public bool SelfContained { get; set; }
    public string TextDensity { get; set; } = string.Empty;
}

public class DeckSequenceListResult : ResultBase
{
    public List<DeckSequenceListItem> Sequences { get; set; } = [];
}

public class DeckSequenceListItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string UsedFor { get; set; } = string.Empty;
    public string Intent { get; set; } = string.Empty;
}

public class DeckSequenceDetailResult : ResultBase
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string UsedFor { get; set; } = string.Empty;
    public string Intent { get; set; } = string.Empty;
    public List<DeckSlideInfo> Slides { get; set; } = [];
}

public class DeckSlideInfo
{
    public string Position { get; set; } = string.Empty;
    public string Purpose { get; set; } = string.Empty;
    public string Archetype { get; set; } = string.Empty;
    public string Density { get; set; } = string.Empty;
}

public class SlidePatternListResult : ResultBase
{
    public string Content { get; set; } = string.Empty;
}

public class IconShapeListResult : ResultBase
{
    public string Content { get; set; } = string.Empty;
}
