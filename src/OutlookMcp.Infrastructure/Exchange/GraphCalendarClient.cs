using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.Infrastructure.Exchange;

/// <summary>
/// Microsoft Graph REST access to the signed-in account's calendar. Requests are bounded,
/// paged, and retried once on throttling. Only event create and delete are exposed, and
/// attendee data is never written, so the API can never send meeting invitations.
/// </summary>
public sealed class GraphCalendarClient(ExchangeTokenProvider tokens, OutlookMcpOptions options, ILogger<GraphCalendarClient> logger) : IExchangeCalendarGateway
{
    private const string GraphBase = "https://graph.microsoft.com/v1.0";
    private const int PageSize = 100;
    private const int MaximumRetries = 3;

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(100) };

    private readonly CalendarSyncOptions _options = options.CalendarSync;

    public Task<ExchangeAuthStatusDto> GetAuthStatusAsync(CancellationToken cancellationToken) => tokens.GetStatusAsync(cancellationToken);

    public Task<ExchangeLoginDto> BeginDeviceCodeLoginAsync(CancellationToken cancellationToken) => tokens.BeginDeviceCodeLoginAsync(cancellationToken);

    public Task<ExchangeAuthStatusDto> LogoutAsync(CancellationToken cancellationToken) => tokens.LogoutAsync(cancellationToken);

    public async Task<ExchangeCalendarDto> GetTargetCalendarAsync(CancellationToken cancellationToken)
    {
        var path = string.IsNullOrWhiteSpace(_options.TargetCalendarId)
            ? "/me/calendar"
            : $"/me/calendars/{Uri.EscapeDataString(_options.TargetCalendarId.Trim())}";
        using var document = await SendForJsonAsync(HttpMethod.Get, path + "?$select=id,name,isDefaultCalendar,owner", null, "read the target calendar", cancellationToken).ConfigureAwait(false);
        var root = document.RootElement;
        var id = GetString(root, "id") ?? throw new OutlookMcpException(ErrorCodes.ExchangeApiFailed,
            "Microsoft Graph returned a calendar without an id.", "Retry; if the problem continues, check CalendarSync.TargetCalendarId.");
        string? owner = null;
        if (root.TryGetProperty("owner", out var ownerNode) && ownerNode.ValueKind == JsonValueKind.Object) owner = GetString(ownerNode, "address");
        var isDefault = root.TryGetProperty("isDefaultCalendar", out var defaultNode) && defaultNode.ValueKind == JsonValueKind.True;
        return new ExchangeCalendarDto(id, GetString(root, "name") ?? string.Empty, owner, isDefault || string.IsNullOrWhiteSpace(_options.TargetCalendarId));
    }

    public async Task<IReadOnlyList<TargetCalendarEvent>> ListEventsAsync(DateTimeOffset windowStart, DateTimeOffset windowEnd, CancellationToken cancellationToken)
    {
        var result = new List<TargetCalendarEvent>();
        var expand = Uri.EscapeDataString($"singleValueExtendedProperties($filter=id eq '{GraphEventMapper.SourceIdPropertyId}' or id eq '{GraphEventMapper.SourceModifiedPropertyId}')");
        var url = $"{CalendarPath()}/calendarView?startDateTime={Uri.EscapeDataString(ToUtcText(windowStart))}&endDateTime={Uri.EscapeDataString(ToUtcText(windowEnd))}" +
            $"&$top={PageSize}&$select=id,subject,start,end,type&$expand={expand}";
        var pages = 0;
        while (url is not null)
        {
            if (++pages > Math.Max(1, _options.MaximumItemsScanned / PageSize) + 1)
            {
                throw new OutlookMcpException(ErrorCodes.ResultLimitExceeded,
                    $"The Exchange calendar returned more than the configured CalendarSync.MaximumItemsScanned limit of {_options.MaximumItemsScanned} events in the sync window.",
                    "Raise CalendarSync.MaximumItemsScanned in config.json or shorten months_ahead, then retry.");
            }

            using var document = await SendForJsonAsync(HttpMethod.Get, url, null, "list Exchange calendar events", cancellationToken).ConfigureAwait(false);
            var root = document.RootElement;
            if (root.TryGetProperty("value", out var values) && values.ValueKind == JsonValueKind.Array)
            {
                foreach (var element in values.EnumerateArray())
                {
                    var parsed = GraphEventMapper.ParseEvent(element);
                    if (parsed is not null) result.Add(parsed);
                }
            }

            url = GetString(root, "@odata.nextLink");
        }

        return result;
    }

    public async Task CreateEventAsync(SourceCalendarOccurrence occurrence, CancellationToken cancellationToken)
    {
        var json = GraphEventMapper.ToEventJson(occurrence);
        using var document = await SendForJsonAsync(HttpMethod.Post, $"{CalendarPath()}/events", json, "create an Exchange calendar event", cancellationToken).ConfigureAwait(false);
    }

    public async Task DeleteEventAsync(string eventId, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(HttpMethod.Delete, $"/me/events/{Uri.EscapeDataString(eventId)}", null, cancellationToken).ConfigureAwait(false);
        if (response.StatusCode is HttpStatusCode.NotFound) return;
        if (!response.IsSuccessStatusCode) throw await ToApiErrorAsync(response, "delete an Exchange calendar event", cancellationToken).ConfigureAwait(false);
    }

    private string CalendarPath() => string.IsNullOrWhiteSpace(_options.TargetCalendarId)
        ? "/me/calendar"
        : $"/me/calendars/{Uri.EscapeDataString(_options.TargetCalendarId.Trim())}";

    private async Task<JsonDocument> SendForJsonAsync(HttpMethod method, string pathOrUrl, string? jsonBody, string operation, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(method, pathOrUrl, jsonBody, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode) throw await ToApiErrorAsync(response, operation, cancellationToken).ConfigureAwait(false);
        var payload = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return JsonDocument.Parse(payload);
        }
        catch (JsonException ex)
        {
            throw new OutlookMcpException(ErrorCodes.ExchangeApiFailed, $"Microsoft Graph returned an unreadable response while trying to {operation}.", "Retry the operation.", ex);
        }
    }

    private async Task<HttpResponseMessage> SendAsync(HttpMethod method, string pathOrUrl, string? jsonBody, CancellationToken cancellationToken)
    {
        for (var attempt = 1; ; attempt++)
        {
            var token = await tokens.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
            using var request = new HttpRequestMessage(method, pathOrUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ? pathOrUrl : GraphBase + pathOrUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            request.Headers.TryAddWithoutValidation("Prefer", "outlook.timezone=\"UTC\"");
            if (jsonBody is not null) request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            var response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false);
            if (response.StatusCode is not (HttpStatusCode.TooManyRequests or HttpStatusCode.ServiceUnavailable) || attempt >= MaximumRetries)
            {
                return response;
            }

            var delay = response.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(2 * attempt);
            response.Dispose();
            logger.LogInformation("Microsoft Graph throttled the request; retrying in {DelaySeconds} s", delay.TotalSeconds);
            await Task.Delay(delay > TimeSpan.FromSeconds(30) ? TimeSpan.FromSeconds(30) : delay, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task<OutlookMcpException> ToApiErrorAsync(HttpResponseMessage response, string operation, CancellationToken cancellationToken)
    {
        var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        var excerpt = body.Length > 600 ? body[..600] : body;
        logger.LogWarning("Microsoft Graph returned {StatusCode} while trying to {Operation}", (int)response.StatusCode, operation);
        if (response.StatusCode is HttpStatusCode.Unauthorized)
        {
            return new OutlookMcpException(ErrorCodes.ExchangeAuthRequired, "Microsoft Graph rejected the sign-in token.",
                "Run outlook_exchange_login again to refresh the sign-in.", new InvalidOperationException(excerpt));
        }

        if (response.StatusCode is HttpStatusCode.Forbidden)
        {
            return new OutlookMcpException(ErrorCodes.ExchangeApiFailed, $"Microsoft Graph refused permission while trying to {operation}.",
                "Verify the app registration has the delegated Calendars.ReadWrite permission and that the signed-in user consented to it.", new InvalidOperationException(excerpt));
        }

        return new OutlookMcpException(ErrorCodes.ExchangeApiFailed, $"Microsoft Graph returned HTTP {(int)response.StatusCode} while trying to {operation}.",
            "Retry; if the problem continues, check the account and configuration.", new InvalidOperationException(excerpt));
    }

    private static string ToUtcText(DateTimeOffset value) => value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

    private static string? GetString(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
}
