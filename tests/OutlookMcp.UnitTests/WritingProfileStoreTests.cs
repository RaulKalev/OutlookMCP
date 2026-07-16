using System.Text.Json;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.WritingStyle;

namespace OutlookMcp.UnitTests;

public sealed class WritingProfileStoreTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "outlook-mcp-profile-tests", Guid.NewGuid().ToString("N"));

    [Fact]
    public async Task SaveUpdateAndRestore_PreservesVersionHistory()
    {
        var store = CreateStore();
        var saved = await store.SaveAsync(ValidProfile("direct"), "initial", "dataset-1", 42, default);
        Assert.True(saved.Exists);

        var updated = await store.UpdateAsync("{\"user_overrides\":[{\"rule\":\"Do not use em dashes.\",\"priority\":100,\"enabled\":true}]}", 3, default);
        Assert.Equal(3, updated.NewMessagesSinceGeneration);
        var version = Assert.Single(await store.ListVersionsAsync(default));
        Assert.Equal(42, version.SourceMessageCount);

        var restored = await store.RestoreAsync(version.VersionId, default);
        var summary = restored.Profile!.Value.GetProperty("summary");
        Assert.Equal("direct", summary.GetProperty("overall_tone").GetString());
    }

    [Theory]
    [InlineData("{\"version\":2,\"summary\":{},\"style_rules\":[]}")]
    [InlineData("{\"version\":1,\"summary\":{},\"style_rules\":[{\"rule\":\"x\",\"confidence\":2}]}")]
    [InlineData("{\"version\":1,\"summary\":{},\"style_rules\":[{\"rule\":\"same\"},{\"rule\":\"same\"}]}")]
    [InlineData("{\"version\":1,\"summary\":{},\"style_rules\":[],\"unsupported\":true}")]
    public async Task Save_RejectsInvalidProfiles(string json)
    {
        var store = CreateStore();
        await Assert.ThrowsAsync<OutlookMcpException>(() => store.SaveAsync(json, null, null, 0, default));
    }

    [Fact]
    public async Task Update_RejectsProtectedGeneratedFields()
    {
        var store = CreateStore();
        await store.SaveAsync(ValidProfile("neutral"), null, null, 1, default);
        await Assert.ThrowsAsync<OutlookMcpException>(() => store.UpdateAsync("{\"version\":2}", 0, default));
    }

    private WritingProfileStore CreateStore()
    {
        Directory.CreateDirectory(_directory);
        var options = new OutlookMcpOptions();
        options.WritingStyle.ProfilePath = Path.Combine(_directory, "profile.json");
        options.WritingStyle.ProfileHistoryPath = Path.Combine(_directory, "history.json");
        return new WritingProfileStore(options);
    }

    private static string ValidProfile(string tone) => JsonSerializer.Serialize(new
    {
        version = 1,
        summary = new { primary_language = "Estonian", overall_tone = tone },
        style_rules = new[] { new { rule = "Get to the point quickly.", confidence = 0.9, evidence_count = 20, source = "automatic" } },
        common_phrases = Array.Empty<object>(),
        contextual_variations = Array.Empty<object>(),
        formatting_preferences = new { },
        avoid_patterns = Array.Empty<object>(),
        user_overrides = Array.Empty<object>()
    });

    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }
}
