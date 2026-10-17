using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.UnitTests;

/// <summary>Gateway whose every member throws; tests override only what they exercise.</summary>
internal abstract class StubOutlookGateway : IOutlookGateway
{
    public virtual ValueTask DisposeAsync() => ValueTask.CompletedTask;
    public virtual Task<OutlookStatusDto> GetStatusAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<StoreDto>> ListStoresAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<FolderDto>> ListFoldersAsync(ListFoldersRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<FolderDto>> FindFoldersAsync(FindFoldersRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<SearchResultDto> SearchEmailsAsync(SearchEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<EmailDetailDto> ReadEmailAsync(ReadEmailRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<BatchReadResultDto> ReadEmailsBatchAsync(ReadEmailsBatchRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<ThreadDto> ReadThreadAsync(ReadThreadRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<SelectionDto> GetSelectedEmailAsync(SelectedEmailRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<RelatedEmailDto>> FindRelatedEmailsAsync(RelatedEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<AttachmentDto>> ListAttachmentsAsync(string messageId, string storeId, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<SavedAttachmentDto> SaveAttachmentAsync(SaveAttachmentRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<AttachmentScratchFileDto> SaveAttachmentToScratchAsync(AttachmentScratchRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<DraftDto> CreateDraftAsync(CreateDraftRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<DraftDto> CreateReplyDraftAsync(CreateReplyDraftRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<DraftDto> CreateForwardDraftAsync(CreateForwardDraftRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<FolderDto> CreateFolderAsync(CreateFolderRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<MoveEmailsResultDto> MoveEmailsAsync(MoveEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<FolderRuleAnalysisDto> AnalyzeFolderForRulesAsync(AnalyzeFolderRulesRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<CreateFolderRuleResultDto> CreateFolderRuleAsync(CreateFolderRuleRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<CalendarFolderDto>> ListCalendarFoldersAsync(string? storeId, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<CalendarOccurrenceReadResult> ReadCalendarOccurrencesAsync(string sourceFolderId, string? sourceStoreId, DateTimeOffset windowStart, DateTimeOffset windowEnd, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<IReadOnlyList<SentFolderDescriptorDto>> DiscoverSentFoldersAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<SentEmailBatchDto> ReadSentFolderBatchAsync(string storeId, string folderId, int startOffset, int batchSize, DateTimeOffset? modifiedSince, CancellationToken cancellationToken) => throw new NotSupportedException();
    public virtual Task<SentEmailReferenceBatchDto> ReadSentFolderReferencesBatchAsync(string storeId, string folderId, int startOffset, int batchSize, CancellationToken cancellationToken) => throw new NotSupportedException();
}
