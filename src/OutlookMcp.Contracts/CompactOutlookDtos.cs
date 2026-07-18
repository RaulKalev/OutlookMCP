using System.Text.Json.Serialization;

namespace OutlookMcp.Contracts;

public sealed record CompactEmailSummaryDto(
    [property: JsonPropertyName("message_id")] string MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("sender_name"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? SenderName,
    [property: JsonPropertyName("sender_email"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? SenderEmail,
    [property: JsonPropertyName("timestamp")] DateTimeOffset Timestamp,
    [property: JsonPropertyName("folder_path")] string FolderPath,
    [property: JsonPropertyName("unread")] bool Unread,
    [property: JsonPropertyName("attachment_filenames")] IReadOnlyList<string> AttachmentFilenames);

public sealed record CompactSearchResultDto(
    [property: JsonPropertyName("messages")] IReadOnlyList<CompactEmailSummaryDto> Messages,
    [property: JsonPropertyName("may_be_incomplete")] bool MayBeIncomplete,
    [property: JsonPropertyName("scope_warning"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? ScopeWarning,
    [property: JsonPropertyName("searched_folder_count")] int SearchedFolderCount,
    [property: JsonPropertyName("scanned_item_count")] int ScannedItemCount,
    [property: JsonPropertyName("scan_truncated")] bool ScanTruncated);

public sealed record CompactEmailDto(
    [property: JsonPropertyName("message_id")] string MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("folder_path")] string FolderPath,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("sender")] EmailAddressDto Sender,
    [property: JsonPropertyName("to")] IReadOnlyList<EmailAddressDto> To,
    [property: JsonPropertyName("cc")] IReadOnlyList<EmailAddressDto> Cc,
    [property: JsonPropertyName("sent_at"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] DateTimeOffset? SentAt,
    [property: JsonPropertyName("received_at"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] DateTimeOffset? ReceivedAt,
    [property: JsonPropertyName("plain_text_body"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? PlainTextBody,
    [property: JsonPropertyName("body_truncated")] bool BodyTruncated,
    [property: JsonPropertyName("original_body_length")] int OriginalBodyLength);

public sealed record CompactBatchEmailResultDto(
    [property: JsonPropertyName("source_message_id")] string SourceMessageId,
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("email"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] CompactEmailDto? Email,
    [property: JsonPropertyName("error"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] ErrorDto? Error);

public sealed record CompactBatchReadResultDto(
    [property: JsonPropertyName("items")] IReadOnlyList<CompactBatchEmailResultDto> Items,
    [property: JsonPropertyName("succeeded_count")] int SucceededCount,
    [property: JsonPropertyName("failed_count")] int FailedCount);

public sealed record CompactSelectionDto(
    [property: JsonPropertyName("messages")] IReadOnlyList<CompactEmailDto> Messages,
    [property: JsonPropertyName("selection_source")] string SelectionSource,
    [property: JsonPropertyName("unsupported_item_count")] int UnsupportedItemCount);

public sealed record CompactRelatedEmailDto(
    [property: JsonPropertyName("message")] CompactEmailSummaryDto Message,
    [property: JsonPropertyName("relevance_reasons")] IReadOnlyList<string> RelevanceReasons,
    [property: JsonPropertyName("score")] int Score);
