using System.Text.Json;
using OutlookMcp.Application.Services;
using OutlookMcp.Infrastructure.Exchange;

namespace OutlookMcp.UnitTests;

public sealed class GraphEventMapperTests
{
    private static readonly DateTimeOffset Start = new(2026, 7, 20, 10, 0, 0, TimeSpan.FromHours(3));

    [Fact]
    public void ToEventJson_MapsAllFieldsWithoutAttendees()
    {
        var occurrence = new SourceCalendarOccurrence(
            "key-1", "stamp-1", "Planning", Start, Start.AddHours(1), false, "FLE Standard Time",
            "Agenda with https://example.com/link", "Room 4", ["Project A", "Internal"], 15, "busy", "private", true, true);
        using var document = JsonDocument.Parse(GraphEventMapper.ToEventJson(occurrence));
        var root = document.RootElement;
        Assert.Equal("Planning", root.GetProperty("subject").GetString());
        Assert.Equal("2026-07-20T10:00:00", root.GetProperty("start").GetProperty("dateTime").GetString());
        Assert.Equal("FLE Standard Time", root.GetProperty("start").GetProperty("timeZone").GetString());
        Assert.Equal("2026-07-20T11:00:00", root.GetProperty("end").GetProperty("dateTime").GetString());
        Assert.False(root.GetProperty("isAllDay").GetBoolean());
        Assert.Equal("busy", root.GetProperty("showAs").GetString());
        Assert.Equal("private", root.GetProperty("sensitivity").GetString());
        Assert.Equal("text", root.GetProperty("body").GetProperty("contentType").GetString());
        Assert.Contains("https://example.com/link", root.GetProperty("body").GetProperty("content").GetString(), StringComparison.Ordinal);
        Assert.Equal("Room 4", root.GetProperty("location").GetProperty("displayName").GetString());
        Assert.Equal(2, root.GetProperty("categories").GetArrayLength());
        Assert.True(root.GetProperty("isReminderOn").GetBoolean());
        Assert.Equal(15, root.GetProperty("reminderMinutesBeforeStart").GetInt32());
        Assert.False(root.GetProperty("responseRequested").GetBoolean());
        Assert.False(root.TryGetProperty("attendees", out _));
        var properties = root.GetProperty("singleValueExtendedProperties");
        Assert.Equal(2, properties.GetArrayLength());
        Assert.Equal("key-1", properties[0].GetProperty("value").GetString());
        Assert.Equal("stamp-1", properties[1].GetProperty("value").GetString());
    }

    [Fact]
    public void ToEventJson_OmitsOptionalFieldsAndReminderWhenAbsent()
    {
        var occurrence = new SourceCalendarOccurrence(
            "key-1", "stamp-1", "Quiet", Start, Start.AddHours(1), true, "UTC",
            null, null, [], null, "free", "normal", false, false);
        using var document = JsonDocument.Parse(GraphEventMapper.ToEventJson(occurrence));
        var root = document.RootElement;
        Assert.True(root.GetProperty("isAllDay").GetBoolean());
        Assert.False(root.GetProperty("isReminderOn").GetBoolean());
        Assert.False(root.TryGetProperty("body", out _));
        Assert.False(root.TryGetProperty("location", out _));
        Assert.False(root.TryGetProperty("categories", out _));
        Assert.False(root.TryGetProperty("reminderMinutesBeforeStart", out _));
    }

    [Fact]
    public void ParseEvent_ReadsSyncTagsAndTimes()
    {
        var json = """
        {
          "id": "evt-1",
          "subject": "Planning",
          "type": "singleInstance",
          "start": { "dateTime": "2026-07-20T07:00:00.0000000", "timeZone": "UTC" },
          "end": { "dateTime": "2026-07-20T08:00:00.0000000", "timeZone": "UTC" },
          "singleValueExtendedProperties": [
            { "id": "String {7c1b3af6-6d7e-4a3a-9e6b-2e5c6a1d9b42} Name OutlookMcpSyncSourceId", "value": "key-1" },
            { "id": "String {7c1b3af6-6d7e-4a3a-9e6b-2e5c6a1d9b42} Name OutlookMcpSyncSourceModified", "value": "stamp-1" }
          ]
        }
        """;
        using var document = JsonDocument.Parse(json);
        var parsed = GraphEventMapper.ParseEvent(document.RootElement);
        Assert.NotNull(parsed);
        Assert.Equal("evt-1", parsed!.EntryId);
        Assert.Equal("key-1", parsed.SyncKey);
        Assert.Equal("stamp-1", parsed.SourceModifiedStamp);
        Assert.False(parsed.IsRecurring);
        Assert.True(parsed.InWindow);
        Assert.Equal(new DateTimeOffset(2026, 7, 20, 7, 0, 0, TimeSpan.Zero), parsed.Start);
    }

    [Fact]
    public void ParseEvent_HandlesMissingTagsAndUnknownType()
    {
        var json = """{ "id": "evt-2", "subject": "Stray", "type": "occurrence" }""";
        using var document = JsonDocument.Parse(json);
        var parsed = GraphEventMapper.ParseEvent(document.RootElement);
        Assert.NotNull(parsed);
        Assert.Null(parsed!.SyncKey);
        Assert.Null(parsed.SourceModifiedStamp);
        Assert.True(parsed.IsRecurring);
    }
}
