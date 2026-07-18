using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Abstractions;

/// <summary>
/// Bounded access to the signed-in Exchange Online account's calendar through the
/// Microsoft Graph API. Sign-in uses the OAuth device-code flow; tokens are cached
/// and refreshed silently so the user authorizes once.
/// </summary>
public interface IExchangeCalendarGateway
{
    Task<ExchangeAuthStatusDto> GetAuthStatusAsync(CancellationToken cancellationToken);
    Task<ExchangeLoginDto> BeginDeviceCodeLoginAsync(CancellationToken cancellationToken);
    Task<ExchangeAuthStatusDto> LogoutAsync(CancellationToken cancellationToken);
    Task<ExchangeCalendarDto> GetTargetCalendarAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<TargetCalendarEvent>> ListEventsAsync(DateTimeOffset windowStart, DateTimeOffset windowEnd, CancellationToken cancellationToken);
    Task CreateEventAsync(SourceCalendarOccurrence occurrence, CancellationToken cancellationToken);
    Task DeleteEventAsync(string eventId, CancellationToken cancellationToken);
}
