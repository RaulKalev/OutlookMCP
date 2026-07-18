namespace OutlookMcp.Application.Services;

/// <summary>A source-calendar event inside the sync window, identified by a stable sync key.</summary>
public sealed record SourceCalendarEvent(
    string EntryId,
    string SyncKey,
    string ModifiedStamp,
    string Subject,
    DateTimeOffset? Start,
    DateTimeOffset? End,
    bool IsRecurring,
    bool RecurrenceExtendsBeyondWindow,
    bool IsMeeting);

/// <summary>A target-calendar event with the sync tags read back from its user properties.</summary>
public sealed record TargetCalendarEvent(
    string EntryId,
    string? SyncKey,
    string? SourceModifiedStamp,
    string Subject,
    DateTimeOffset? Start,
    DateTimeOffset? End,
    bool IsRecurring,
    bool InWindow);

public sealed record CalendarSyncUpdate(SourceCalendarEvent Source, TargetCalendarEvent Target);

public sealed record CalendarSyncPlan(
    IReadOnlyList<SourceCalendarEvent> Adds,
    IReadOnlyList<CalendarSyncUpdate> Updates,
    IReadOnlyList<TargetCalendarEvent> Deletes,
    int UnchangedCount);

/// <summary>
/// Computes a one-way calendar sync plan. The source calendar is never modified; the plan
/// only ever adds to, refreshes, or prunes the dedicated target calendar. Target events are
/// matched by the sync key stamped onto each copy; unmatched target events are deleted only
/// when they fall inside the sync window, so history outside the window is left alone.
/// </summary>
public static class CalendarSyncPlanner
{
    public static CalendarSyncPlan Plan(IReadOnlyList<SourceCalendarEvent> sourceEvents, IReadOnlyList<TargetCalendarEvent> targetEvents)
    {
        var adds = new List<SourceCalendarEvent>();
        var updates = new List<CalendarSyncUpdate>();
        var deletes = new List<TargetCalendarEvent>();
        var unchanged = 0;
        var candidatesByKey = targetEvents
            .Where(value => !string.IsNullOrWhiteSpace(value.SyncKey))
            .GroupBy(value => value.SyncKey!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => new Queue<TargetCalendarEvent>(group), StringComparer.Ordinal);
        var matchedEntryIds = new HashSet<string>(StringComparer.Ordinal);

        foreach (var source in sourceEvents)
        {
            if (!candidatesByKey.TryGetValue(source.SyncKey, out var candidates) || candidates.Count == 0)
            {
                adds.Add(source);
                continue;
            }

            var target = candidates.Dequeue();
            matchedEntryIds.Add(target.EntryId);
            if (string.Equals(target.SourceModifiedStamp, source.ModifiedStamp, StringComparison.Ordinal)) unchanged++;
            else updates.Add(new CalendarSyncUpdate(source, target));
        }

        foreach (var target in targetEvents)
        {
            if (!matchedEntryIds.Contains(target.EntryId) && target.InWindow) deletes.Add(target);
        }

        return new CalendarSyncPlan(adds, updates, deletes, unchanged);
    }

    /// <summary>
    /// Decides whether an event belongs to the sync window. Single events count when they
    /// overlap the window; recurring series count while the series is still active anywhere
    /// in the window, because the series master carries every occurrence with it.
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
