using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Services;
using OutlookMcp.Application.WritingStyle;
using OutlookMcp.Contracts;
using OutlookMcp.Infrastructure.WritingStyle;

namespace OutlookMcp.UnitTests;

public sealed class WritingStyleCoordinatorTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "outlook-mcp-coordinator-tests", Guid.NewGuid().ToString("N"));

    [Fact]
    public async Task Scan_ResumesAcrossInvocationsWithoutTotalMessageCap()
    {
        using var fixture = CreateFixture();
        var first = await fixture.Coordinator.ScanAsync(2, 1, false, true, default);
        Assert.False(first.Complete);
        Assert.Equal(2, first.ProcessedInThisRun);
        Assert.Equal(1, first.Remaining);

        var second = await fixture.Coordinator.ScanAsync(2, 1, false, true, default);
        Assert.True(second.Complete);
        Assert.Equal(1, second.ProcessedInThisRun);
        var status = await fixture.Coordinator.GetStatusAsync(default);
        Assert.Equal(3, status.TotalProcessed);
        Assert.Equal(0, status.TotalRemaining);
        Assert.True(status.InitialScanComplete);
    }

    [Fact]
    public async Task Retrieval_BoostsMetadataDeduplicatesAndExcludesCurrentMessage()
    {
        using var fixture = CreateFixture();
        await fixture.Coordinator.ScanAsync(10, 1, false, true, default);
        var examples = await fixture.Coordinator.FindExamplesAsync(new StyleExampleQueryDto(
            DraftSubject: "ABC-123 kooskõlastus", RecipientAddresses: ["kaur@example.com"], ProjectKeywords: ["ABC-123"],
            CommunicationIntent: "requesting_approval", MaxResults: 3, MaximumCharactersPerExample: 500), default);

        var first = Assert.Single(examples);
        Assert.True(first.SameRecipient);
        Assert.True(first.SameProject);
        Assert.True(first.SameCommunicationIntent);
        Assert.Equal("untrusted_data", first.TrustLevel);

        var excluded = await fixture.Coordinator.FindExamplesAsync(new StyleExampleQueryDto(CurrentMessageId: first.MessageId, CurrentStoreId: "store",
            DraftSubject: "ABC-123", ProjectKeywords: ["ABC-123"], MaxResults: 3, MaximumCharactersPerExample: 500), default);
        Assert.DoesNotContain(excluded, value => value.MessageId == first.MessageId);
    }

    [Fact]
    public async Task DraftContext_PrioritisesExplicitInstructionThenUserOverride()
    {
        using var fixture = CreateFixture();
        await fixture.Coordinator.ScanAsync(10, 1, false, true, default);
        const string profile = """
            {"version":1,"summary":{"primary_language":"Estonian"},"style_rules":[{"rule":"Generated rule","confidence":0.8}],"user_overrides":[{"rule":"Confirmed override","priority":100,"enabled":true}]}
            """;
        await fixture.Coordinator.SaveProfileAsync(profile, null, "dataset", default);

        var context = await fixture.Coordinator.PrepareDraftContextAsync(new StyleExampleQueryDto(DraftSubject: "ABC-123", DraftContext: "Keep this reply short.", ProjectKeywords: ["ABC-123"], MaxResults: 2), 5_000, false, default);
        Assert.StartsWith("Current drafting instruction:", context.DraftingRules[0], StringComparison.Ordinal);
        Assert.Equal("Confirmed override", context.DraftingRules[1]);
        Assert.Equal("Generated rule", context.DraftingRules[2]);
        Assert.All(context.RelevantExamples, example => Assert.Equal("untrusted_data", example.TrustLevel));
        Assert.Contains(context.SafetyNotes, note => note.Contains("unsent", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Sync_ReconcilesMissingSentEntryIdsWithoutReadingBodiesAgain()
    {
        using var fixture = CreateFixture();
        await fixture.Coordinator.ScanAsync(10, 1, false, true, default);
        fixture.Gateway.MissingEntryId = "two";

        var result = await fixture.Coordinator.SyncAsync(default);
        Assert.Equal(1, result.MissingMessagesDetected);
        Assert.Empty(result.Errors);
    }

    private Fixture CreateFixture()
    {
        Directory.CreateDirectory(_directory);
        var options = new OutlookMcpOptions();
        options.WritingStyle.DatabasePath = Path.Combine(_directory, "index.db");
        options.WritingStyle.ProfilePath = Path.Combine(_directory, "profile.json");
        options.WritingStyle.ProfileHistoryPath = Path.Combine(_directory, "history.json");
        options.WritingStyle.SyncBeforeDraftContext = false;
        var repository = new SqliteStyleIndexRepository(options, NullLogger<SqliteStyleIndexRepository>.Instance);
        var gateway = new FakeGateway();
        var coordinator = new WritingStyleCoordinator(gateway, repository, new AuthoredTextExtractor(new EmailBodyCleaner()), new CommunicationIntentClassifier(),
            new WritingProfileStore(options), options, NullLogger<WritingStyleCoordinator>.Instance);
        return new Fixture(coordinator, repository, gateway);
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }

    private sealed record Fixture(WritingStyleCoordinator Coordinator, SqliteStyleIndexRepository Repository, FakeGateway Gateway) : IDisposable
    {
        public void Dispose() { Coordinator.Dispose(); Repository.Dispose(); }
    }

    private sealed class FakeGateway : IOutlookGateway
    {
        private readonly SentFolderDescriptorDto _folder = new("store", "Test", "sent", "\\Test\\Sent", 3, "test");
        private readonly SentEmailSourceDto[] _messages =
        [
            Message("one", "ABC-123 kooskõlastus", "Tere Kaur\n\nPalun kinnita ABC-123 lahendus.\n\nLugupidamisega\nRaul", "kaur@example.com"),
            Message("two", "Koosolek", "Tere\n\nKohtume homme kell 10.\n\nLugupidamisega\nRaul", "mari@example.com"),
            Message("three", "Re: ABC-123 kooskõlastus", "Tere Kaur\n\nPalun kinnita ABC-123 lahendus.\n\nFrom: Kaur\nVana kiri", "kaur@example.com")
        ];
        public string? MissingEntryId { get; set; }

        public Task<IReadOnlyList<SentFolderDescriptorDto>> DiscoverSentFoldersAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<SentFolderDescriptorDto>>([_folder]);
        public Task<SentEmailBatchDto> ReadSentFolderBatchAsync(string storeId, string folderId, int startOffset, int batchSize, DateTimeOffset? modifiedSince, CancellationToken cancellationToken)
        {
            var items = _messages.Skip(startOffset).Take(batchSize).ToArray();
            var next = startOffset + items.Length;
            return Task.FromResult(new SentEmailBatchDto(_folder, startOffset, next, _messages.Length, items, next >= _messages.Length));
        }

        public Task<SentEmailReferenceBatchDto> ReadSentFolderReferencesBatchAsync(string storeId, string folderId, int startOffset, int batchSize, CancellationToken cancellationToken)
        {
            var current = _messages.Where(value => !string.Equals(value.EntryId, MissingEntryId, StringComparison.Ordinal)).ToArray();
            var entryIds = current.Skip(startOffset).Take(batchSize).Select(value => value.EntryId).ToArray();
            var next = startOffset + entryIds.Length;
            return Task.FromResult(new SentEmailReferenceBatchDto(_folder with { TotalItems = current.Length }, startOffset, next, current.Length, entryIds, next >= current.Length));
        }

        private static SentEmailSourceDto Message(string id, string subject, string body, string recipient) => new(id, "message-" + id, "store", "sent", "\\Test\\Sent", null, null, subject,
            subject, DateTimeOffset.UtcNow, DateTimeOffset.UtcNow, "Raul", "raul@example.com", [new(null, recipient, recipient)], [], [], body, null, [], "successfully_processed", null);

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
        public Task<OutlookStatusDto> GetStatusAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<StoreDto>> ListStoresAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<FolderDto>> ListFoldersAsync(ListFoldersRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<FolderDto>> FindFoldersAsync(FindFoldersRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<SearchResultDto> SearchEmailsAsync(SearchEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<EmailDetailDto> ReadEmailAsync(ReadEmailRequest request, CancellationToken cancellationToken) => Task.FromResult(new EmailDetailDto(
            request.MessageId, request.StoreId, "sent", "\\Test\\Sent", "ABC-123 kooskõlastus", new("Kaur", "kaur@example.com", "kaur@example.com"),
            [new("Raul", "raul@example.com", "raul@example.com")], [], [], DateTimeOffset.UtcNow, DateTimeOffset.UtcNow, DateTimeOffset.UtcNow, DateTimeOffset.UtcNow,
            "normal", false, null, [], null, null, null, "ABC-123", null, "Palun kinnita ABC-123.", null, false, 24, 24, [], "untrusted"));
        public Task<BatchReadResultDto> ReadEmailsBatchAsync(ReadEmailsBatchRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<ThreadDto> ReadThreadAsync(ReadThreadRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<SelectionDto> GetSelectedEmailAsync(SelectedEmailRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<RelatedEmailDto>> FindRelatedEmailsAsync(RelatedEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<AttachmentDto>> ListAttachmentsAsync(string messageId, string storeId, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<SavedAttachmentDto> SaveAttachmentAsync(SaveAttachmentRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<DraftDto> CreateDraftAsync(CreateDraftRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<DraftDto> CreateReplyDraftAsync(CreateReplyDraftRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<DraftDto> CreateForwardDraftAsync(CreateForwardDraftRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<FolderDto> CreateFolderAsync(CreateFolderRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<MoveEmailsResultDto> MoveEmailsAsync(MoveEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<FolderRuleAnalysisDto> AnalyzeFolderForRulesAsync(AnalyzeFolderRulesRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<CreateFolderRuleResultDto> CreateFolderRuleAsync(CreateFolderRuleRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<CalendarFolderDto>> ListCalendarFoldersAsync(string? storeId, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<CalendarSyncResultDto> SyncCalendarAsync(SyncCalendarRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
    }
}
