using OutlookMcp.Contracts;

namespace OutlookMcp.Server.Mcp;

internal static class CompactDtoMapper
{
    public static CompactSearchResultDto ToCompact(SearchResultDto value) => new(
        value.Messages.Select(ToCompact).ToArray(),
        value.MayBeIncomplete,
        value.ScopeWarning,
        value.SearchedFolderCount,
        value.ScannedItemCount,
        value.ScanTruncated);

    public static CompactBatchReadResultDto ToCompact(BatchReadResultDto value) => new(
        value.Items.Select(item => new CompactBatchEmailResultDto(
            item.SourceMessageId,
            item.Success,
            item.Email is null ? null : ToCompact(item.Email),
            item.Error)).ToArray(),
        value.SucceededCount,
        value.FailedCount);

    public static CompactSelectionDto ToCompact(SelectionDto value) => new(
        value.Messages.Select(ToCompact).ToArray(),
        value.SelectionSource,
        value.UnsupportedItemCount);

    public static IReadOnlyList<CompactRelatedEmailDto> ToCompact(IReadOnlyList<RelatedEmailDto> values) =>
        values.Select(value => new CompactRelatedEmailDto(ToCompact(value.Message), value.RelevanceReasons, value.Score)).ToArray();

    private static CompactEmailSummaryDto ToCompact(EmailSummaryDto value) => new(
        value.MessageId,
        value.StoreId,
        value.Subject,
        value.SenderName,
        value.SenderEmail,
        value.Timestamp,
        value.FolderPath,
        value.Unread,
        value.AttachmentFilenames);

    private static CompactEmailDto ToCompact(EmailDetailDto value) => new(
        value.MessageId,
        value.StoreId,
        value.FolderPath,
        value.Subject,
        value.Sender,
        value.To,
        value.Cc,
        value.SentAt,
        value.ReceivedAt,
        value.PlainTextBody,
        value.BodyTruncated,
        value.OriginalBodyLength);
}
