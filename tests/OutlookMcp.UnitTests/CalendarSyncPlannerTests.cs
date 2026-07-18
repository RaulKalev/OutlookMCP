using OutlookMcp.Application.Services;

namespace OutlookMcp.UnitTests;

public sealed class CalendarSyncPlannerTests
{
    private static readonly DateTimeOffset WindowStart = new(2026, 7, 18, 0, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset WindowEnd = WindowStart.AddMonths(3);

    [Fact]
    public void Plan_AddsSourceEventsMissingFromTarget()
    {
        var plan = CalendarSyncPlanner.Plan([Source("s1", "key-1"), Source("s2", "key-2")], []);
        Assert.Equal(2, plan.Adds.Count);
        Assert.Empty(plan.Updates);
        Assert.Empty(plan.Deletes);
        Assert.Equal(0, plan.UnchangedCount);
    }

    [Fact]
    public void Plan_LeavesMatchingEventsWithSameStampUnchanged()
    {
        var plan = CalendarSyncPlanner.Plan([Source("s1", "key-1", "stamp-a")], [Target("t1", "key-1", "stamp-a")]);
        Assert.Empty(plan.Adds);
        Assert.Empty(plan.Updates);
        Assert.Empty(plan.Deletes);
        Assert.Equal(1, plan.UnchangedCount);
    }

    [Fact]
    public void Plan_UpdatesWhenSourceModifiedStampDiffers()
    {
        var plan = CalendarSyncPlanner.Plan([Source("s1", "key-1", "stamp-new")], [Target("t1", "key-1", "stamp-old")]);
        var update = Assert.Single(plan.Updates);
        Assert.Equal("s1", update.Source.EntryId);
        Assert.Equal("t1", update.Target.EntryId);
        Assert.Empty(plan.Adds);
        Assert.Empty(plan.Deletes);
    }

    [Fact]
    public void Plan_DeletesInWindowTargetsRemovedFromSource()
    {
        var plan = CalendarSyncPlanner.Plan([], [Target("t1", "key-1", "stamp-a")]);
        var delete = Assert.Single(plan.Deletes);
        Assert.Equal("t1", delete.EntryId);
    }

    [Fact]
    public void Plan_LeavesOutOfWindowTargetsAlone()
    {
        var pastEvent = Target("t1", "key-old", "stamp-a") with { InWindow = false };
        var plan = CalendarSyncPlanner.Plan([], [pastEvent]);
        Assert.Empty(plan.Deletes);
    }

    [Fact]
    public void Plan_DeletesUntaggedInWindowTargets()
    {
        var stray = Target("t1", null, null);
        var plan = CalendarSyncPlanner.Plan([Source("s1", "key-1")], [stray, Target("t2", "key-1", "stamp")]);
        var delete = Assert.Single(plan.Deletes);
        Assert.Equal("t1", delete.EntryId);
        Assert.Null(delete.SyncKey);
    }

    [Fact]
    public void Plan_DeletesDuplicateTargetCopiesWithSameKey()
    {
        var plan = CalendarSyncPlanner.Plan(
            [Source("s1", "key-1", "stamp-a")],
            [Target("t1", "key-1", "stamp-a"), Target("t2", "key-1", "stamp-a")]);
        Assert.Equal(1, plan.UnchangedCount);
        var delete = Assert.Single(plan.Deletes);
        Assert.Equal("t2", delete.EntryId);
    }

    [Fact]
    public void Plan_MatchesOutOfWindowTargetSoMovedSourceEventIsNotDuplicated()
    {
        var movedCopy = Target("t1", "key-1", "stamp-old") with { InWindow = false };
        var plan = CalendarSyncPlanner.Plan([Source("s1", "key-1", "stamp-new")], [movedCopy]);
        var update = Assert.Single(plan.Updates);
        Assert.Equal("t1", update.Target.EntryId);
        Assert.Empty(plan.Adds);
    }

    [Fact]
    public void Plan_DuplicateSourceKeysEachConsumeOneTargetCopy()
    {
        var plan = CalendarSyncPlanner.Plan(
            [Source("s1", "key-1", "stamp-a"), Source("s2", "key-1", "stamp-a")],
            [Target("t1", "key-1", "stamp-a")]);
        Assert.Equal(1, plan.UnchangedCount);
        var add = Assert.Single(plan.Adds);
        Assert.Equal("s2", add.EntryId);
        Assert.Empty(plan.Deletes);
    }

    [Theory]
    [InlineData(1, 2, true)]
    [InlineData(-10, -5, false)]
    [InlineData(-1, 1, true)]
    [InlineData(100, 101, false)]
    public void IsEventInWindow_SingleEventsUseOverlap(int startOffsetDays, int endOffsetDays, bool expected)
    {
        var inWindow = CalendarSyncPlanner.IsEventInWindow(
            WindowStart.AddDays(startOffsetDays), WindowStart.AddDays(endOffsetDays),
            false, false, null, WindowStart, WindowEnd);
        Assert.Equal(expected, inWindow);
    }

    [Fact]
    public void IsEventInWindow_SingleEventWithoutStartIsExcluded()
    {
        Assert.False(CalendarSyncPlanner.IsEventInWindow(null, null, false, false, null, WindowStart, WindowEnd));
    }

    [Fact]
    public void IsEventInWindow_OpenEndedOldSeriesIsIncluded()
    {
        Assert.True(CalendarSyncPlanner.IsEventInWindow(
            WindowStart.AddYears(-2), WindowStart.AddYears(-2).AddHours(1),
            true, false, null, WindowStart, WindowEnd));
    }

    [Fact]
    public void IsEventInWindow_SeriesEndedBeforeWindowIsExcluded()
    {
        Assert.False(CalendarSyncPlanner.IsEventInWindow(
            WindowStart.AddYears(-2), WindowStart.AddYears(-2).AddHours(1),
            true, true, WindowStart.AddDays(-1), WindowStart, WindowEnd));
    }

    [Fact]
    public void IsEventInWindow_SeriesEndingInsideWindowIsIncluded()
    {
        Assert.True(CalendarSyncPlanner.IsEventInWindow(
            WindowStart.AddYears(-1), WindowStart.AddYears(-1).AddHours(1),
            true, true, WindowStart.AddDays(10), WindowStart, WindowEnd));
    }

    [Fact]
    public void IsEventInWindow_SeriesStartingAfterWindowIsExcluded()
    {
        Assert.False(CalendarSyncPlanner.IsEventInWindow(
            WindowEnd.AddDays(1), WindowEnd.AddDays(1).AddHours(1),
            true, false, null, WindowStart, WindowEnd));
    }

    private static SourceCalendarEvent Source(string entryId, string syncKey, string stamp = "stamp") =>
        new(entryId, syncKey, stamp, "Subject " + entryId, WindowStart.AddDays(1), WindowStart.AddDays(1).AddHours(1), false, false, false);

    private static TargetCalendarEvent Target(string entryId, string? syncKey, string? stamp) =>
        new(entryId, syncKey, stamp, "Subject " + entryId, WindowStart.AddDays(1), WindowStart.AddDays(1).AddHours(1), false, true);
}
