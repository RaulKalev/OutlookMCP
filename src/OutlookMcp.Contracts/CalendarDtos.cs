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

public sealed record CalendarSyncItemDto(
    [property: JsonPropertyName("action")] string Action,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("start")] DateTimeOffset? Start,
    [property: JsonPropertyName("end")] DateTimeOffset? End,
    [property: JsonPropertyName("is_recurring")] bool IsRecurring,
    [property: JsonPropertyName("applied")] bool Applied,
    [property: JsonPropertyName("error"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] ErrorDto? Error);

public sealed record CalendarSyncResultDto(
    [property: JsonPropertyName("source_calendar")] FolderDto SourceCalendar,
    [property: JsonPropertyName("target_calendar")] FolderDto TargetCalendar,
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
