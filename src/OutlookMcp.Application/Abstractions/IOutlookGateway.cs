using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Abstractions;

public interface IOutlookGateway : IAsyncDisposable
{
    Task<OutlookStatusDto> GetStatusAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<StoreDto>> ListStoresAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<FolderDto>> ListFoldersAsync(ListFoldersRequest request, CancellationToken cancellationToken);
    Task<IReadOnlyList<FolderDto>> FindFoldersAsync(FindFoldersRequest request, CancellationToken cancellationToken);
    Task<SearchResultDto> SearchEmailsAsync(SearchEmailsRequest request, CancellationToken cancellationToken);
    Task<EmailDetailDto> ReadEmailAsync(ReadEmailRequest request, CancellationToken cancellationToken);
    Task<BatchReadResultDto> ReadEmailsBatchAsync(ReadEmailsBatchRequest request, CancellationToken cancellationToken);
    Task<ThreadDto> ReadThreadAsync(ReadThreadRequest request, CancellationToken cancellationToken);
    Task<SelectionDto> GetSelectedEmailAsync(SelectedEmailRequest request, CancellationToken cancellationToken);
    Task<IReadOnlyList<RelatedEmailDto>> FindRelatedEmailsAsync(RelatedEmailsRequest request, CancellationToken cancellationToken);
    Task<IReadOnlyList<AttachmentDto>> ListAttachmentsAsync(string messageId, string storeId, CancellationToken cancellationToken);
    Task<SavedAttachmentDto> SaveAttachmentAsync(SaveAttachmentRequest request, CancellationToken cancellationToken);
    Task<DraftDto> CreateDraftAsync(CreateDraftRequest request, CancellationToken cancellationToken);
    Task<DraftDto> CreateReplyDraftAsync(CreateReplyDraftRequest request, CancellationToken cancellationToken);
    Task<DraftDto> CreateForwardDraftAsync(CreateForwardDraftRequest request, CancellationToken cancellationToken);
    Task<FolderDto> CreateFolderAsync(CreateFolderRequest request, CancellationToken cancellationToken);
    Task<MoveEmailsResultDto> MoveEmailsAsync(MoveEmailsRequest request, CancellationToken cancellationToken);
    Task<FolderRuleAnalysisDto> AnalyzeFolderForRulesAsync(AnalyzeFolderRulesRequest request, CancellationToken cancellationToken);
    Task<CreateFolderRuleResultDto> CreateFolderRuleAsync(CreateFolderRuleRequest request, CancellationToken cancellationToken);
    Task<IReadOnlyList<CalendarFolderDto>> ListCalendarFoldersAsync(string? storeId, CancellationToken cancellationToken);
    Task<CalendarSyncResultDto> SyncCalendarAsync(SyncCalendarRequest request, CancellationToken cancellationToken);
    Task<IReadOnlyList<SentFolderDescriptorDto>> DiscoverSentFoldersAsync(CancellationToken cancellationToken);
    Task<SentEmailBatchDto> ReadSentFolderBatchAsync(string storeId, string folderId, int startOffset, int batchSize, DateTimeOffset? modifiedSince, CancellationToken cancellationToken);
    Task<SentEmailReferenceBatchDto> ReadSentFolderReferencesBatchAsync(string storeId, string folderId, int startOffset, int batchSize, CancellationToken cancellationToken);
}
