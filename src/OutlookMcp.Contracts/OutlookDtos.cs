using System.Text.Json.Serialization;

namespace OutlookMcp.Contracts;

public sealed record OutlookItemReference(string EntryId, string StoreId);

public sealed record OutlookStatusDto(
    [property: JsonPropertyName("outlook_classic_installed")] bool OutlookClassicInstalled,
    [property: JsonPropertyName("outlook_running")] bool OutlookRunning,
    [property: JsonPropertyName("mapi_available")] bool MapiAvailable,
    [property: JsonPropertyName("outlook_version")] string? OutlookVersion,
    [property: JsonPropertyName("profile_name")] string? ProfileName,
    [property: JsonPropertyName("store_count")] int StoreCount,
    [property: JsonPropertyName("server_version")] string ServerVersion,
    [property: JsonPropertyName("warnings")] IReadOnlyList<string> Warnings);

public sealed record StoreDto(
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("display_name")] string DisplayName,
    [property: JsonPropertyName("store_type")] string? StoreType,
    [property: JsonPropertyName("is_default")] bool IsDefault,
    [property: JsonPropertyName("is_accessible")] bool IsAccessible,
    [property: JsonPropertyName("root_folder_name")] string? RootFolderName);

public sealed record FolderDto(
    [property: JsonPropertyName("folder_id")] string FolderId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("display_name")] string DisplayName,
    [property: JsonPropertyName("full_path")] string FullPath,
    [property: JsonPropertyName("folder_type")] string FolderType,
    [property: JsonPropertyName("unread_count")] int? UnreadCount,
    [property: JsonPropertyName("total_item_count")] int? TotalItemCount,
    [property: JsonPropertyName("contains_mail_items")] bool ContainsMailItems,
    [property: JsonPropertyName("child_folder_count")] int ChildFolderCount);

public sealed record EmailAddressDto(
    [property: JsonPropertyName("display_name")] string? DisplayName,
    [property: JsonPropertyName("address")] string? Address,
    [property: JsonPropertyName("raw_address")] string? RawAddress);

public sealed record AttachmentDto(
    [property: JsonPropertyName("attachment_id")] int AttachmentId,
    [property: JsonPropertyName("filename")] string Filename,
    [property: JsonPropertyName("mime_type")] string? MimeType,
    [property: JsonPropertyName("size")] long Size,
    [property: JsonPropertyName("inline")] bool Inline,
    [property: JsonPropertyName("content_id")] string? ContentId);

public sealed record EmailSummaryDto(
    [property: JsonPropertyName("message_id")] string MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("folder_id")] string FolderId,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("sender_name")] string? SenderName,
    [property: JsonPropertyName("sender_email")] string? SenderEmail,
    [property: JsonPropertyName("recipients_summary")] string RecipientsSummary,
    [property: JsonPropertyName("timestamp")] DateTimeOffset Timestamp,
    [property: JsonPropertyName("folder_path")] string FolderPath,
    [property: JsonPropertyName("unread")] bool Unread,
    [property: JsonPropertyName("attachment_count")] int AttachmentCount,
    [property: JsonPropertyName("attachment_filenames")] IReadOnlyList<string> AttachmentFilenames,
    [property: JsonPropertyName("body_preview")] string? BodyPreview,
    [property: JsonPropertyName("conversation_topic")] string? ConversationTopic,
    [property: JsonPropertyName("conversation_id")] string? ConversationId,
    [property: JsonPropertyName("external_content_warning"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? ExternalContentWarning);

public sealed record SearchResultDto(
    [property: JsonPropertyName("messages")] IReadOnlyList<EmailSummaryDto> Messages,
    [property: JsonPropertyName("may_be_incomplete")] bool MayBeIncomplete,
    [property: JsonPropertyName("scope_warning")] string? ScopeWarning,
    [property: JsonPropertyName("searched_folder_count")] int SearchedFolderCount,
    [property: JsonPropertyName("scanned_item_count")] int ScannedItemCount,
    [property: JsonPropertyName("scan_truncated")] bool ScanTruncated,
    [property: JsonPropertyName("external_content_warning")] string ExternalContentWarning);

public sealed record EmailDetailDto(
    [property: JsonPropertyName("message_id")] string MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("folder_id")] string FolderId,
    [property: JsonPropertyName("folder_path")] string FolderPath,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("sender")] EmailAddressDto Sender,
    [property: JsonPropertyName("to")] IReadOnlyList<EmailAddressDto> To,
    [property: JsonPropertyName("cc")] IReadOnlyList<EmailAddressDto> Cc,
    [property: JsonPropertyName("bcc")] IReadOnlyList<EmailAddressDto> Bcc,
    [property: JsonPropertyName("sent_at")] DateTimeOffset? SentAt,
    [property: JsonPropertyName("received_at")] DateTimeOffset? ReceivedAt,
    [property: JsonPropertyName("created_at")] DateTimeOffset? CreatedAt,
    [property: JsonPropertyName("modified_at")] DateTimeOffset? ModifiedAt,
    [property: JsonPropertyName("importance")] string Importance,
    [property: JsonPropertyName("unread")] bool Unread,
    [property: JsonPropertyName("flag_state")] string? FlagState,
    [property: JsonPropertyName("categories")] IReadOnlyList<string> Categories,
    [property: JsonPropertyName("internet_message_id")] string? InternetMessageId,
    [property: JsonPropertyName("in_reply_to")] string? InReplyTo,
    [property: JsonPropertyName("references")] string? References,
    [property: JsonPropertyName("conversation_topic")] string? ConversationTopic,
    [property: JsonPropertyName("conversation_id")] string? ConversationId,
    [property: JsonPropertyName("plain_text_body")] string? PlainTextBody,
    [property: JsonPropertyName("html_body")] string? HtmlBody,
    [property: JsonPropertyName("body_truncated")] bool BodyTruncated,
    [property: JsonPropertyName("original_body_length")] int OriginalBodyLength,
    [property: JsonPropertyName("returned_body_length")] int ReturnedBodyLength,
    [property: JsonPropertyName("attachments")] IReadOnlyList<AttachmentDto> Attachments,
    [property: JsonPropertyName("external_content_warning")] string ExternalContentWarning);

public sealed record ThreadMessageDto(
    [property: JsonPropertyName("message_id")] string MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("folder")] string Folder,
    [property: JsonPropertyName("sender")] EmailAddressDto Sender,
    [property: JsonPropertyName("recipients")] IReadOnlyList<EmailAddressDto> Recipients,
    [property: JsonPropertyName("timestamp")] DateTimeOffset Timestamp,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("plain_text_body")] string PlainTextBody,
    [property: JsonPropertyName("attachment_filenames")] IReadOnlyList<string> AttachmentFilenames);

public sealed record ThreadDto(
    [property: JsonPropertyName("messages")] IReadOnlyList<ThreadMessageDto> Messages,
    [property: JsonPropertyName("assembly_method")] string AssemblyMethod,
    [property: JsonPropertyName("confidence")] string Confidence,
    [property: JsonPropertyName("external_content_warning")] string ExternalContentWarning);

public sealed record RelatedEmailDto(
    [property: JsonPropertyName("message")] EmailSummaryDto Message,
    [property: JsonPropertyName("relevance_reasons")] IReadOnlyList<string> RelevanceReasons,
    [property: JsonPropertyName("score")] int Score);

public sealed record SelectionDto(
    [property: JsonPropertyName("messages")] IReadOnlyList<EmailDetailDto> Messages,
    [property: JsonPropertyName("selection_source")] string SelectionSource,
    [property: JsonPropertyName("unsupported_item_count")] int UnsupportedItemCount);

public sealed record DraftDto(
    [property: JsonPropertyName("message_id")] string MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("folder_id")] string FolderId,
    [property: JsonPropertyName("folder_path")] string FolderPath,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("to")] IReadOnlyList<EmailAddressDto> To,
    [property: JsonPropertyName("cc")] IReadOnlyList<EmailAddressDto> Cc,
    [property: JsonPropertyName("bcc")] IReadOnlyList<EmailAddressDto> Bcc,
    [property: JsonPropertyName("attachment_filenames")] IReadOnlyList<string> AttachmentFilenames,
    [property: JsonPropertyName("sent")] bool Sent);

public sealed record SavedAttachmentDto(
    [property: JsonPropertyName("path")] string Path,
    [property: JsonPropertyName("filename")] string Filename,
    [property: JsonPropertyName("size")] long Size,
    [property: JsonPropertyName("warning")] string Warning);

public sealed record BatchEmailResultDto(
    [property: JsonPropertyName("source_message_id")] string SourceMessageId,
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("email")] EmailDetailDto? Email,
    [property: JsonPropertyName("error")] ErrorDto? Error);

public sealed record BatchReadResultDto(
    [property: JsonPropertyName("items")] IReadOnlyList<BatchEmailResultDto> Items,
    [property: JsonPropertyName("requested_count")] int RequestedCount,
    [property: JsonPropertyName("succeeded_count")] int SucceededCount,
    [property: JsonPropertyName("failed_count")] int FailedCount);

public sealed record MoveEmailResultDto(
    [property: JsonPropertyName("source_message_id")] string SourceMessageId,
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("moved")] bool Moved,
    [property: JsonPropertyName("message_id")] string? MessageId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("subject")] string? Subject,
    [property: JsonPropertyName("folder_id")] string? FolderId,
    [property: JsonPropertyName("folder_path")] string? FolderPath,
    [property: JsonPropertyName("error")] ErrorDto? Error);

public sealed record MoveEmailsResultDto(
    [property: JsonPropertyName("destination_folder")] FolderDto DestinationFolder,
    [property: JsonPropertyName("dry_run")] bool DryRun,
    [property: JsonPropertyName("items")] IReadOnlyList<MoveEmailResultDto> Items,
    [property: JsonPropertyName("requested_count")] int RequestedCount,
    [property: JsonPropertyName("succeeded_count")] int SucceededCount,
    [property: JsonPropertyName("failed_count")] int FailedCount);

public sealed record OutlookDiagnosticDto(
    [property: JsonPropertyName("draft_folder_accessible")] bool DraftFolderAccessible,
    [property: JsonPropertyName("selected_item_accessible")] bool SelectedItemAccessible);
