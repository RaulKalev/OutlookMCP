using System.Text.Json.Serialization;

namespace OutlookMcp.Contracts;

public sealed record CalendarFolderDto(
    [property: JsonPropertyName("folder_id")] string FolderId,
    [property: JsonPropertyName("store_id")] string StoreId,
    [property: JsonPropertyName("store_name")] string StoreName,
    [property: JsonPropertyName("display_name")] string DisplayName,
    [property: JsonPropertyName("full_path")] string FullPath,
    [property: JsonPropertyName("is_default_calendar")] bool IsDefaultCalendar,
    [property: JsonPropertyName("total_item_count")] int? TotalItemCount);

public sealed record ExchangeCalendarDto(
    [property: JsonPropertyName("calendar_id")] string CalendarId,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("owner")] string? Owner,
    [property: JsonPropertyName("is_default_calendar")] bool IsDefaultCalendar);

public sealed record ExchangeLoginDto(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("account"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Account,
    [property: JsonPropertyName("verification_url"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? VerificationUrl,
    [property: JsonPropertyName("user_code"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? UserCode,
    [property: JsonPropertyName("expires_at"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] DateTimeOffset? ExpiresAt,
    [property: JsonPropertyName("message")] string Message);

public sealed record ExchangeAuthStatusDto(
    [property: JsonPropertyName("state")] string State,
    [property: JsonPropertyName("account"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Account,
    [property: JsonPropertyName("token_cache_persisted")] bool TokenCachePersisted,
    [property: JsonPropertyName("message")] string Message);

public sealed record CalendarSyncItemDto(
    [property: JsonPropertyName("action")] string Action,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("start")] DateTimeOffset? Start,
    [property: JsonPropertyName("end")] DateTimeOffset? End,
    [property: JsonPropertyName("from_recurring_series")] bool FromRecurringSeries,
    [property: JsonPropertyName("applied")] bool Applied,
    [property: JsonPropertyName("error"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] ErrorDto? Error);

public sealed record CalendarSyncResultDto(
    [property: JsonPropertyName("source_calendar")] FolderDto SourceCalendar,
    [property: JsonPropertyName("target_calendar")] ExchangeCalendarDto TargetCalendar,
    [property: JsonPropertyName("window_start")] DateTimeOffset WindowStart,
    [property: JsonPropertyName("window_end")] DateTimeOffset WindowEnd,
    [property: JsonPropertyName("months_ahead")] int MonthsAhead,
    [property: JsonPropertyName("dry_run")] bool DryRun,
    [property: JsonPropertyName("source_event_count")] int SourceEventCount,
    [property: JsonPropertyName("added_count")] int AddedCount,
    [property: JsonPropertyName("updated_count")] int UpdatedCount,
    [property: JsonPropertyName("deleted_count")] int DeletedCount,
    [property: JsonPropertyName("unchanged_count")] int UnchangedCount,
    [property: JsonPropertyName("failed_count")] int FailedCount,
    [property: JsonPropertyName("items")] IReadOnlyList<CalendarSyncItemDto> Items,
    [property: JsonPropertyName("warnings")] IReadOnlyList<string> Warnings);
