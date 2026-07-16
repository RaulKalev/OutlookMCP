using OutlookMcp.Contracts;

namespace OutlookMcp.Application.WritingStyle;

public interface IStyleIndexRepository
{
    string DatabasePath { get; }
    Task InitializeAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<StyleScanCheckpointDto>> GetCheckpointsAsync(CancellationToken cancellationToken);
    Task UpsertCheckpointAsync(StyleScanCheckpointDto checkpoint, CancellationToken cancellationToken);
    Task<(int Inserted, int Updated)> UpsertMessagesAsync(IReadOnlyList<IndexedSentEmailDto> messages, CancellationToken cancellationToken);
    Task<StyleRepositoryCountsDto> GetCountsAsync(CancellationToken cancellationToken);
    Task<string?> GetStateAsync(string key, CancellationToken cancellationToken);
    Task SetStateAsync(string key, string value, CancellationToken cancellationToken);
    Task RebuildRecurringBlocksAsync(CancellationToken cancellationToken);
    Task<StyleStatisticsDto> GetStatisticsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<StyleDatasetExampleDto>> GetRepresentativeExamplesAsync(int maximumExamples, int maximumCharactersPerExample, CancellationToken cancellationToken);
    Task<IReadOnlyDictionary<string, int>> GetCommonPhrasesAsync(int maximumPhrases, CancellationToken cancellationToken);
    Task<IReadOnlyList<StyleSearchCandidateDto>> SearchExamplesAsync(string query, int candidateLimit, CancellationToken cancellationToken);
    Task<IReadOnlySet<string>> GetEntryIdsAsync(string storeId, string folderId, CancellationToken cancellationToken);
    Task ReplaceMissingEntryIdsAsync(string storeId, string folderId, IReadOnlyCollection<string> entryIds, CancellationToken cancellationToken);
}
