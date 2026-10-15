using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

/// <summary>
/// One expanded source-calendar occurrence inside the sync window, carrying every field
/// needed to recreate it as a standalone Exchange event. Recurring series are expanded
/// occurrence by occurrence, so cancelled or moved single occurrences follow automatically.
/// </summary>
public sealed record SourceCalendarOccurrence(
    string SyncKey,
    string ModifiedStamp,
    string Subject,
    DateTimeOffset? Start,
    DateTimeOffset? End,
    bool IsAllDay,
    string TimeZoneId,
    string? BodyText,
    string? Location,
    IReadOnlyList<string> Categories,
    int? ReminderMinutesBeforeStart,
    string BusyStatus,
    string Sensitivity,
    bool IsMeeting,
    bool FromRecurringSeries);

/// <summary>A target-calendar event with the sync tags read back from its extended properties.</summary>
public sealed record TargetCalendarEvent(
    string EntryId,
    string? SyncKey,
    string? SourceModifiedStamp,
    string Subject,
    DateTimeOffset? Start,
    DateTimeOffset? End,
    bool IsRecurring,
    bool InWindow,
    bool IsAllDay = false);

public sealed record CalendarSyncUpdate(SourceCalendarOccurrence Source, TargetCalendarEvent Target);

public sealed record CalendarSyncPlan(
    IReadOnlyList<SourceCalendarOccurrence> Adds,
    IReadOnlyList<CalendarSyncUpdate> Updates,
    IReadOnlyList<TargetCalendarEvent> Deletes,
    int UnchangedCount,
    int MatchedByContentCount = 0);

/// <summary>The bounded outcome of expanding one local calendar's occurrences for the sync window.</summary>
public sealed record CalendarOccurrenceReadResult(
    FolderDto SourceCalendar,
    IReadOnlyList<SourceCalendarOccurrence> Occurrences,
    int SkippedCount);

/// <summary>
/// Computes a one-way calendar sync plan. The source calendar is never modified; the plan
/// only ever adds to, refreshes, or prunes the sync-owned target calendar. Target events are
/// matched by the sync key stamped onto each copy first; source occurrences left over are then
/// matched to any remaining target event with the same subject, start, and end, so copies that
/// already exist in Exchange (created by hand, by an invitation, or under an older key) are
/// kept instead of being duplicated and deleted. Unmatched target events are deleted only
/// when they fall inside the sync window, so history outside the window is left alone.
/// </summary>
public static class CalendarSyncPlanner
{
    public static CalendarSyncPlan Plan(IReadOnlyList<SourceCalendarOccurrence> sourceEvents, IReadOnlyList<TargetCalendarEvent> targetEvents)
    {
        var adds = new List<SourceCalendarOccurrence>();
        var updates = new List<CalendarSyncUpdate>();
        var deletes = new List<TargetCalendarEvent>();
        var unchanged = 0;
        var candidatesByKey = targetEvents
            .Where(value => !string.IsNullOrWhiteSpace(value.SyncKey))
            .GroupBy(value => value.SyncKey!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => new Queue<TargetCalendarEvent>(group), StringComparer.Ordinal);
        var matchedEntryIds = new HashSet<string>(StringComparer.Ordinal);
        var unmatchedSources = new List<SourceCalendarOccurrence>();

        foreach (var source in sourceEvents)
        {
            if (!candidatesByKey.TryGetValue(source.SyncKey, out var candidates) || candidates.Count == 0)
            {
                unmatchedSources.Add(source);
                continue;
            }

            var target = candidates.Dequeue();
            matchedEntryIds.Add(target.EntryId);
            if (string.Equals(target.SourceModifiedStamp, source.ModifiedStamp, StringComparison.Ordinal)) unchanged++;
            else updates.Add(new CalendarSyncUpdate(source, target));
        }

        var matchedByContent = 0;
        var candidatesByContent = targetEvents
            .Where(value => !matchedEntryIds.Contains(value.EntryId) && value.Start is not null)
            .GroupBy(value => ContentKey(value.Subject, value.Start, value.End, value.IsAllDay), StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => new Queue<TargetCalendarEvent>(group), StringComparer.Ordinal);
        foreach (var source in unmatchedSources)
        {
            if (source.Start is not null
                && candidatesByContent.TryGetValue(ContentKey(source.Subject, source.Start, source.End, source.IsAllDay), out var candidates)
                && candidates.Count > 0)
            {
                matchedEntryIds.Add(candidates.Dequeue().EntryId);
                matchedByContent++;
                unchanged++;
            }
            else
            {
                adds.Add(source);
            }
        }

        foreach (var target in targetEvents)
        {
            if (!matchedEntryIds.Contains(target.EntryId) && target.InWindow) deletes.Add(target);
        }

        return new CalendarSyncPlan(adds, updates, deletes, unchanged, matchedByContent);
    }

    /// <summary>
    /// Identity used to recognise an existing target copy that carries no matching sync tag.
    /// Timed events compare instants to the minute; all-day events compare calendar dates,
    /// because Graph reports them as floating midnights rather than zone-converted instants.
    /// </summary>
    private static string ContentKey(string subject, DateTimeOffset? start, DateTimeOffset? end, bool isAllDay)
    {
        var effectiveEnd = end ?? start;
        var times = isAllDay
            ? $"D|{start!.Value.DateTime:yyyy-MM-dd}|{effectiveEnd!.Value.DateTime:yyyy-MM-dd}"
            : $"T|{start!.Value.UtcDateTime:yyyy-MM-ddTHH:mm}|{effectiveEnd!.Value.UtcDateTime:yyyy-MM-ddTHH:mm}";
        return times + "|" + SubjectKey(subject);
    }

    private static string SubjectKey(string subject) =>
        string.Join(' ', subject.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToUpperInvariant();

    /// <summary>
    /// Decides whether an event belongs to the sync window. Single events count when they
    /// overlap the window; recurring series count while the series is still active anywhere
    /// in the window.
    /// </summary>
    public static bool IsEventInWindow(
        DateTimeOffset? start,
        DateTimeOffset? end,
        bool isRecurring,
        bool recurrenceHasEnd,
        DateTimeOffset? recurrenceEnd,
        DateTimeOffset windowStart,
        DateTimeOffset windowEnd)
    {
        if (isRecurring)
        {
            var seriesStarted = start is null || start <= windowEnd;
            var seriesStillActive = !recurrenceHasEnd || recurrenceEnd is null || recurrenceEnd >= windowStart;
            return seriesStarted && seriesStillActive;
        }

        if (start is null) return false;
        var effectiveEnd = end ?? start;
        return effectiveEnd >= windowStart && start <= windowEnd;
    }
}
