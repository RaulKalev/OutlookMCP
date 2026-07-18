using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.UnitTests;

public sealed class CalendarSyncCoordinatorTests
{
    private static readonly DateTimeOffset Start = new(DateTime.Today);

    [Fact]
    public async Task Sync_RequiresSignedInAccount()
    {
        var exchange = new FakeExchangeGateway { AuthState = "signed_out" };
        var coordinator = Coordinator(new FakeOutlookGateway(), exchange);
        var exception = await Assert.ThrowsAsync<OutlookMcpException>(() => coordinator.SyncAsync(new("folder"), CancellationToken.None));
        Assert.Equal(ErrorCodes.ExchangeAuthRequired, exception.Code);
    }

    [Fact]
    public async Task Sync_RequiresSourceCalendar()
    {
        var coordinator = Coordinator(new FakeOutlookGateway(), new FakeExchangeGateway());
        var exception = await Assert.ThrowsAsync<OutlookMcpException>(() => coordinator.SyncAsync(new(), CancellationToken.None));
        Assert.Equal(ErrorCodes.InvalidArgument, exception.Code);
    }

    [Fact]
    public async Task Sync_RejectsMonthsBeyondConfiguredMaximum()
    {
        var coordinator = Coordinator(new FakeOutlookGateway(), new FakeExchangeGateway());
        var exception = await Assert.ThrowsAsync<OutlookMcpException>(() => coordinator.SyncAsync(new("folder", MonthsAhead: 99), CancellationToken.None));
        Assert.Equal(ErrorCodes.InvalidArgument, exception.Code);
    }

    [Fact]
    public async Task DryRun_PlansWithoutCallingGraphWrites()
    {
        var outlook = new FakeOutlookGateway { Occurrences = [Occurrence("key-new")] };
        var exchange = new FakeExchangeGateway { Events = [TargetEvent("evt-1", "key-old")] };
        var coordinator = Coordinator(outlook, exchange);
        var result = await coordinator.SyncAsync(new("folder"), CancellationToken.None);
        Assert.True(result.DryRun);
        Assert.Equal(1, result.AddedCount);
        Assert.Equal(1, result.DeletedCount);
        Assert.Equal(0, result.FailedCount);
        Assert.All(result.Items, item => Assert.False(item.Applied));
        Assert.Empty(exchange.CreatedKeys);
        Assert.Empty(exchange.DeletedIds);
    }

    [Fact]
    public async Task Apply_ExecutesDeletesUpdatesAndAdds()
    {
        var outlook = new FakeOutlookGateway { Occurrences = [Occurrence("key-new"), Occurrence("key-changed", "stamp-b")] };
        var exchange = new FakeExchangeGateway
        {
            Events = [TargetEvent("evt-orphan", "key-gone"), TargetEvent("evt-changed", "key-changed", "stamp-a")]
        };
        var coordinator = Coordinator(outlook, exchange);
        var result = await coordinator.SyncAsync(new("folder", DryRun: false), CancellationToken.None);
        Assert.Equal(1, result.AddedCount);
        Assert.Equal(1, result.UpdatedCount);
        Assert.Equal(1, result.DeletedCount);
        Assert.Equal(0, result.FailedCount);
        Assert.Equal(["evt-orphan", "evt-changed"], exchange.DeletedIds);
        Assert.Equal(["key-changed", "key-new"], exchange.CreatedKeys);
        Assert.All(result.Items, item => Assert.True(item.Applied));
    }

    [Fact]
    public async Task Apply_ReportsPerItemFailuresAndContinues()
    {
        var outlook = new FakeOutlookGateway { Occurrences = [Occurrence("key-fail"), Occurrence("key-ok")] };
        var exchange = new FakeExchangeGateway { FailingCreateKey = "key-fail" };
        var coordinator = Coordinator(outlook, exchange);
        var result = await coordinator.SyncAsync(new("folder", DryRun: false), CancellationToken.None);
        Assert.Equal(1, result.FailedCount);
        Assert.Contains("key-ok", exchange.CreatedKeys);
        var failure = Assert.Single(result.Items, item => item.Error is not null);
        Assert.False(failure.Applied);
        Assert.Equal(ErrorCodes.ExchangeApiFailed, failure.Error!.Code);
    }

    [Fact]
    public async Task Sync_WarnsAboutSyncOwnedTargetWithAccount()
    {
        var coordinator = Coordinator(new FakeOutlookGateway(), new FakeExchangeGateway());
        var result = await coordinator.SyncAsync(new("folder"), CancellationToken.None);
        Assert.Contains(result.Warnings, warning => warning.Contains("user@example.com", StringComparison.Ordinal) && warning.Contains("sync-owned", StringComparison.Ordinal));
    }

    private static CalendarSyncCoordinator Coordinator(FakeOutlookGateway outlook, FakeExchangeGateway exchange) =>
        new(outlook, exchange, new OutlookMcpOptions(), NullLogger<CalendarSyncCoordinator>.Instance);

    private static SourceCalendarOccurrence Occurrence(string key, string stamp = "stamp") =>
        new(key, stamp, "Subject " + key, Start.AddDays(1), Start.AddDays(1).AddHours(1),
            false, "FLE Standard Time", "body", "location", [], 15, "busy", "normal", false, false);

    private static TargetCalendarEvent TargetEvent(string id, string? key, string stamp = "stamp") =>
        new(id, key, stamp, "Subject " + id, Start.AddDays(1), Start.AddDays(1).AddHours(1), false, true);

    private sealed class FakeExchangeGateway : IExchangeCalendarGateway
    {
        public string AuthState { get; set; } = "signed_in";
        public List<TargetCalendarEvent> Events { get; set; } = [];
        public string? FailingCreateKey { get; set; }
        public List<string> CreatedKeys { get; } = [];
        public List<string> DeletedIds { get; } = [];

        public Task<ExchangeAuthStatusDto> GetAuthStatusAsync(CancellationToken cancellationToken) =>
            Task.FromResult(new ExchangeAuthStatusDto(AuthState, AuthState == "signed_in" ? "user@example.com" : null, true, "test"));

        public Task<ExchangeLoginDto> BeginDeviceCodeLoginAsync(CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<ExchangeAuthStatusDto> LogoutAsync(CancellationToken cancellationToken) => throw new NotSupportedException();

        public Task<ExchangeCalendarDto> GetTargetCalendarAsync(CancellationToken cancellationToken) =>
            Task.FromResult(new ExchangeCalendarDto("cal-1", "Calendar", "user@example.com", true));

        public Task<IReadOnlyList<TargetCalendarEvent>> ListEventsAsync(DateTimeOffset windowStart, DateTimeOffset windowEnd, CancellationToken cancellationToken) =>
            Task.FromResult<IReadOnlyList<TargetCalendarEvent>>(Events);

        public Task CreateEventAsync(SourceCalendarOccurrence occurrence, CancellationToken cancellationToken)
        {
            if (string.Equals(occurrence.SyncKey, FailingCreateKey, StringComparison.Ordinal))
            {
                throw new OutlookMcpException(ErrorCodes.ExchangeApiFailed, "Simulated Graph failure.", "Retry.");
            }

            CreatedKeys.Add(occurrence.SyncKey);
            return Task.CompletedTask;
        }

        public Task DeleteEventAsync(string eventId, CancellationToken cancellationToken)
        {
            DeletedIds.Add(eventId);
            return Task.CompletedTask;
        }
    }

    private sealed class FakeOutlookGateway : IOutlookGateway
    {
        public List<SourceCalendarOccurrence> Occurrences { get; set; } = [];
        public int SkippedCount { get; set; }

        public Task<CalendarOccurrenceReadResult> ReadCalendarOccurrencesAsync(string sourceFolderId, string? sourceStoreId, DateTimeOffset windowStart, DateTimeOffset windowEnd, CancellationToken cancellationToken) =>
            Task.FromResult(new CalendarOccurrenceReadResult(
                new FolderDto(sourceFolderId, "store-1", "Calendar", "\\\\Store\\Calendar", "appointment", null, Occurrences.Count, false, 0),
                Occurrences, SkippedCount));

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
        public Task<OutlookStatusDto> GetStatusAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<StoreDto>> ListStoresAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<FolderDto>> ListFoldersAsync(ListFoldersRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<IReadOnlyList<FolderDto>> FindFoldersAsync(FindFoldersRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<SearchResultDto> SearchEmailsAsync(SearchEmailsRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<EmailDetailDto> ReadEmailAsync(ReadEmailRequest request, CancellationToken cancellationToken) => throw new NotSupportedException();
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
        public Task<IReadOnlyList<SentFolderDescriptorDto>> DiscoverSentFoldersAsync(CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<SentEmailBatchDto> ReadSentFolderBatchAsync(string storeId, string folderId, int startOffset, int batchSize, DateTimeOffset? modifiedSince, CancellationToken cancellationToken) => throw new NotSupportedException();
        public Task<SentEmailReferenceBatchDto> ReadSentFolderReferencesBatchAsync(string storeId, string folderId, int startOffset, int batchSize, CancellationToken cancellationToken) => throw new NotSupportedException();
    }
}
