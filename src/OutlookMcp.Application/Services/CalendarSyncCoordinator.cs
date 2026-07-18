using Microsoft.Extensions.Logging;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

/// <summary>
/// Orchestrates the one-way calendar sync: expanded occurrences are read from the local
/// Outlook source calendar, compared against the signed-in Exchange account's calendar,
/// and the resulting plan is applied through the Graph API. The local calendar is never
/// modified and each applied action reports its own success or error.
/// </summary>
public sealed class CalendarSyncCoordinator(
    IOutlookGateway outlook,
    IExchangeCalendarGateway exchange,
    OutlookMcpOptions options,
    ILogger<CalendarSyncCoordinator> logger)
{
    public async Task<CalendarSyncResultDto> SyncAsync(SyncCalendarRequest request, CancellationToken cancellationToken)
    {
        var sync = options.CalendarSync;
        var months = request.MonthsAhead ?? sync.DefaultMonthsAhead;
        if (months < 1 || months > sync.MaximumMonthsAhead)
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, $"months_ahead must be between 1 and {sync.MaximumMonthsAhead}.", "Correct the request parameters and retry.");
        }

        var sourceFolderId = FirstConfigured(request.SourceCalendarFolderId, sync.SourceCalendarFolderId)
            ?? throw new OutlookMcpException(ErrorCodes.InvalidArgument,
                "source_calendar_folder_id is required.",
                "Run outlook_list_calendars and pass the source calendar's folder_id, or set CalendarSync.SourceCalendarFolderId in config.json.");
        var sourceStoreId = FirstConfigured(request.SourceStoreId, sync.SourceStoreId);
        var windowStart = new DateTimeOffset(DateTime.Today);
        var windowEnd = windowStart.AddMonths(months);

        var authStatus = await exchange.GetAuthStatusAsync(cancellationToken).ConfigureAwait(false);
        if (!string.Equals(authStatus.State, "signed_in", StringComparison.Ordinal))
        {
            throw new OutlookMcpException(ErrorCodes.ExchangeAuthRequired, "No Exchange account is signed in.", "Run outlook_exchange_login, complete the device-code sign-in, then retry the sync.");
        }

        var sourceRead = await outlook.ReadCalendarOccurrencesAsync(sourceFolderId, sourceStoreId, windowStart, windowEnd, cancellationToken).ConfigureAwait(false);
        var targetCalendar = await exchange.GetTargetCalendarAsync(cancellationToken).ConfigureAwait(false);
        var targetEvents = await exchange.ListEventsAsync(windowStart, windowEnd, cancellationToken).ConfigureAwait(false);
        var plan = CalendarSyncPlanner.Plan(sourceRead.Occurrences, targetEvents);
        var warnings = BuildWarnings(authStatus.Account, targetCalendar, plan, sourceRead.SkippedCount);
        var items = new List<CalendarSyncItemDto>(plan.Deletes.Count + plan.Updates.Count + plan.Adds.Count);
        var failed = 0;

        if (request.DryRun)
        {
            foreach (var target in plan.Deletes) items.Add(new CalendarSyncItemDto("delete", target.Subject, target.Start, target.End, target.IsRecurring, false, null));
            foreach (var update in plan.Updates) items.Add(new CalendarSyncItemDto("update", update.Source.Subject, update.Source.Start, update.Source.End, update.Source.FromRecurringSeries, false, null));
            foreach (var source in plan.Adds) items.Add(new CalendarSyncItemDto("add", source.Subject, source.Start, source.End, source.FromRecurringSeries, false, null));
        }
        else
        {
            foreach (var target in plan.Deletes)
            {
                items.Add(await ApplyAsync("delete", target.Subject, target.Start, target.End, target.IsRecurring,
                    () => exchange.DeleteEventAsync(target.EntryId, cancellationToken), () => failed++).ConfigureAwait(false));
            }

            foreach (var update in plan.Updates)
            {
                items.Add(await ApplyAsync("update", update.Source.Subject, update.Source.Start, update.Source.End, update.Source.FromRecurringSeries,
                    async () =>
                    {
                        await exchange.DeleteEventAsync(update.Target.EntryId, cancellationToken).ConfigureAwait(false);
                        await exchange.CreateEventAsync(update.Source, cancellationToken).ConfigureAwait(false);
                    }, () => failed++).ConfigureAwait(false));
            }

            foreach (var source in plan.Adds)
            {
                items.Add(await ApplyAsync("add", source.Subject, source.Start, source.End, source.FromRecurringSeries,
                    () => exchange.CreateEventAsync(source, cancellationToken), () => failed++).ConfigureAwait(false));
            }
        }

        logger.LogInformation(
            "Calendar sync completed; DryRun={DryRun}, SourceEvents={SourceEvents}, Adds={Adds}, Updates={Updates}, Deletes={Deletes}, Unchanged={Unchanged}, Failed={Failed}",
            request.DryRun, sourceRead.Occurrences.Count, plan.Adds.Count, plan.Updates.Count, plan.Deletes.Count, plan.UnchangedCount, failed);
        return new CalendarSyncResultDto(sourceRead.SourceCalendar, targetCalendar, windowStart, windowEnd, months, request.DryRun,
            sourceRead.Occurrences.Count, plan.Adds.Count, plan.Updates.Count, plan.Deletes.Count, plan.UnchangedCount, failed, items, warnings);
    }

    private async Task<CalendarSyncItemDto> ApplyAsync(string action, string subject, DateTimeOffset? start, DateTimeOffset? end, bool fromSeries, Func<Task> operation, Action onFailure)
    {
        try
        {
            await operation().ConfigureAwait(false);
            return new CalendarSyncItemDto(action, subject, start, end, fromSeries, true, null);
        }
        catch (OutlookMcpException ex) when (ex.Code is not (ErrorCodes.ExchangeAuthRequired or ErrorCodes.ExchangeNotConfigured or ErrorCodes.OperationCancelled))
        {
            onFailure();
            logger.LogWarning("Calendar sync {Action} failed with {ErrorCode}", action, ex.Code);
            return new CalendarSyncItemDto(action, subject, start, end, fromSeries, false, ex.ToError(options.Logging.IncludeTechnicalDetails));
        }
    }

    private static IReadOnlyList<string> BuildWarnings(string? account, ExchangeCalendarDto targetCalendar, CalendarSyncPlan plan, int sourceSkipped)
    {
        var warnings = new List<string>
        {
            $"Signed in as {account ?? "an unknown account"}. The Exchange calendar '{targetCalendar.Name}' is treated as sync-owned: window events missing from the local source calendar are deleted there. The local calendar is never modified."
        };
        var copied = plan.Adds.Concat(plan.Updates.Select(value => value.Source)).ToArray();
        var fromSeries = copied.Count(value => value.FromRecurringSeries);
        if (fromSeries > 0) warnings.Add($"{fromSeries} copied events are occurrences of recurring series, expanded into individual events within the sync window.");
        var meetings = copied.Count(value => value.IsMeeting);
        if (meetings > 0) warnings.Add($"{meetings} copied events are meetings; they are copied as plain calendar data without attendee lists, so no invitations, responses, or cancellations can ever be sent.");
        var untaggedDeletes = plan.Deletes.Count(value => string.IsNullOrWhiteSpace(value.SyncKey));
        if (untaggedDeletes > 0) warnings.Add($"{untaggedDeletes} Exchange events were not created by this sync and fall inside the window, so they are planned for deletion. Review the dry-run item list before applying.");
        if (sourceSkipped > 0) warnings.Add($"{sourceSkipped} source calendar items could not be read and were skipped; their Exchange copies may be deleted until they become readable again.");
        return warnings;
    }

    private static string? FirstConfigured(params string?[] values) => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim();
}
