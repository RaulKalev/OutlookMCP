using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using OutlookMcp.Application.Services;

namespace OutlookMcp.Infrastructure.Exchange;

/// <summary>
/// Pure mapping between expanded local calendar occurrences and Microsoft Graph event
/// resources. Attendees are never mapped, so the Graph API can never send invitations
/// on behalf of the sync.
/// </summary>
public static class GraphEventMapper
{
    private const string PropertyNamespace = "{7c1b3af6-6d7e-4a3a-9e6b-2e5c6a1d9b42}";
    public const string SourceIdPropertyId = $"String {PropertyNamespace} Name OutlookMcpSyncSourceId";
    public const string SourceModifiedPropertyId = $"String {PropertyNamespace} Name OutlookMcpSyncSourceModified";

    public static string ToEventJson(SourceCalendarOccurrence occurrence)
    {
        var start = occurrence.Start ?? throw new ArgumentException("An occurrence without a start time cannot be mapped.", nameof(occurrence));
        var end = occurrence.End ?? start;
        var payload = new JsonObject
        {
            ["subject"] = occurrence.Subject,
            ["start"] = TimeNode(start, occurrence.TimeZoneId),
            ["end"] = TimeNode(end, occurrence.TimeZoneId),
            ["isAllDay"] = occurrence.IsAllDay,
            ["showAs"] = occurrence.BusyStatus,
            ["sensitivity"] = occurrence.Sensitivity,
            ["isReminderOn"] = occurrence.ReminderMinutesBeforeStart is not null,
            ["responseRequested"] = false,
            ["singleValueExtendedProperties"] = new JsonArray(
                new JsonObject { ["id"] = SourceIdPropertyId, ["value"] = occurrence.SyncKey },
                new JsonObject { ["id"] = SourceModifiedPropertyId, ["value"] = occurrence.ModifiedStamp })
        };
        if (!string.IsNullOrWhiteSpace(occurrence.BodyText))
        {
            payload["body"] = new JsonObject { ["contentType"] = "text", ["content"] = occurrence.BodyText };
        }

        if (!string.IsNullOrWhiteSpace(occurrence.Location))
        {
            payload["location"] = new JsonObject { ["displayName"] = occurrence.Location };
        }

        if (occurrence.Categories.Count > 0)
        {
            payload["categories"] = new JsonArray(occurrence.Categories.Select(value => (JsonNode)value).ToArray());
        }

        if (occurrence.ReminderMinutesBeforeStart is not null)
        {
            payload["reminderMinutesBeforeStart"] = occurrence.ReminderMinutesBeforeStart.Value;
        }

        return payload.ToJsonString();
    }

    public static TargetCalendarEvent? ParseEvent(JsonElement element)
    {
        var id = GetString(element, "id");
        if (string.IsNullOrWhiteSpace(id)) return null;
        string? syncKey = null;
        string? sourceModified = null;
        if (element.TryGetProperty("singleValueExtendedProperties", out var properties) && properties.ValueKind == JsonValueKind.Array)
        {
            foreach (var property in properties.EnumerateArray())
            {
                var propertyId = GetString(property, "id");
                var value = GetString(property, "value");
                if (EndsWithName(propertyId, "OutlookMcpSyncSourceId")) syncKey = value;
                else if (EndsWithName(propertyId, "OutlookMcpSyncSourceModified")) sourceModified = value;
            }
        }

        var type = GetString(element, "type");
        return new TargetCalendarEvent(
            id!,
            syncKey,
            sourceModified,
            GetString(element, "subject") ?? string.Empty,
            ParseTime(element, "start"),
            ParseTime(element, "end"),
            type is not null && !string.Equals(type, "singleInstance", StringComparison.OrdinalIgnoreCase),
            true);
    }

    private static JsonObject TimeNode(DateTimeOffset value, string timeZoneId) => new()
    {
        ["dateTime"] = value.DateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
        ["timeZone"] = timeZoneId
    };

    private static DateTimeOffset? ParseTime(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out var node) || node.ValueKind != JsonValueKind.Object) return null;
        var text = GetString(node, "dateTime");
        if (text is null || !DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)) return null;
        var zone = GetString(node, "timeZone");
        return string.Equals(zone, "UTC", StringComparison.OrdinalIgnoreCase)
            ? new DateTimeOffset(DateTime.SpecifyKind(parsed, DateTimeKind.Utc))
            : new DateTimeOffset(DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified), TimeSpan.Zero);
    }

    private static bool EndsWithName(string? propertyId, string name) =>
        propertyId is not null && propertyId.EndsWith("Name " + name, StringComparison.OrdinalIgnoreCase);

    private static string? GetString(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
}
