using System.Text.Json;
using System.Text.Json.Nodes;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.WritingStyle;

public sealed class WritingProfileStore
{
    private static readonly HashSet<string> AllowedProfileFields = new(StringComparer.Ordinal)
    {
        "version", "generated_at", "source_statistics", "summary", "style_rules", "common_phrases",
        "contextual_variations", "formatting_preferences", "avoid_patterns", "user_overrides",
        "preferred_phrases", "forbidden_phrases", "generation_metadata"
    };
    private static readonly HashSet<string> AllowedPatchFields = new(AllowedProfileFields.Where(value => value is not "version" and not "generated_at" and not "source_statistics" and not "generation_metadata"), StringComparer.Ordinal);
    private readonly WritingStyleOptions _options;
    private readonly JsonSerializerOptions _jsonOptions = new() { WriteIndented = true };

    public WritingProfileStore(OutlookMcpOptions options) => _options = options.WritingStyle;
    public string ProfilePath => AppPaths.ExpandPath(_options.ProfilePath);
    public string HistoryPath => AppPaths.ExpandPath(_options.ProfileHistoryPath);

    public async Task<WritingProfileResultDto> GetAsync(int newMessagesSinceGeneration, CancellationToken cancellationToken)
    {
        if (!File.Exists(ProfilePath)) return new(false, null, ProfilePath, null, null, null, false, newMessagesSinceGeneration);
        var node = JsonNode.Parse(await File.ReadAllTextAsync(ProfilePath, cancellationToken).ConfigureAwait(false))?.AsObject() ?? throw Invalid("The writing profile file is invalid JSON.");
        var metadata = node["generation_metadata"] as JsonObject;
        var generatedAt = TryDate(node["generated_at"]?.GetValue<string>());
        var element = JsonSerializer.SerializeToElement(node, _jsonOptions);
        return new(true, element, ProfilePath, generatedAt, metadata?["source_dataset_id"]?.GetValue<string>(), metadata?["notes"]?.GetValue<string>(), newMessagesSinceGeneration >= _options.ProfileRefreshRecommendationThreshold, newMessagesSinceGeneration);
    }

    public async Task<WritingProfileResultDto> SaveAsync(string profileJson, string? notes, string? sourceDatasetId, int sourceMessageCount, CancellationToken cancellationToken)
    {
        var profile = ParseAndValidate(profileJson);
        profile["generated_at"] = DateTimeOffset.UtcNow.ToString("O");
        profile["generation_metadata"] = new JsonObject
        {
            ["source_dataset_id"] = sourceDatasetId,
            ["notes"] = Truncate(notes, 4_000),
            ["generation_method"] = "ai_assisted_from_local_representative_dataset",
            ["source_message_count"] = sourceMessageCount
        };
        await ArchiveCurrentAsync(cancellationToken).ConfigureAwait(false);
        await WriteProfileAsync(profile, cancellationToken).ConfigureAwait(false);
        return await GetAsync(0, cancellationToken).ConfigureAwait(false);
    }

    public async Task<WritingProfileResultDto> UpdateAsync(string patchJson, int newMessagesSinceGeneration, CancellationToken cancellationToken)
    {
        if (!File.Exists(ProfilePath)) throw Invalid("No writing profile exists to update.");
        var patch = JsonNode.Parse(patchJson)?.AsObject() ?? throw Invalid("patch_json must be a JSON object.");
        foreach (var key in patch.Select(value => value.Key)) if (!AllowedPatchFields.Contains(key)) throw Invalid($"Unsupported profile update field: {key}");
        var current = JsonNode.Parse(await File.ReadAllTextAsync(ProfilePath, cancellationToken).ConfigureAwait(false))!.AsObject();
        foreach (var (key, value) in patch) current[key] = value?.DeepClone();
        current["generated_at"] = DateTimeOffset.UtcNow.ToString("O");
        Validate(current);
        await ArchiveCurrentAsync(cancellationToken).ConfigureAwait(false);
        await WriteProfileAsync(current, cancellationToken).ConfigureAwait(false);
        return await GetAsync(newMessagesSinceGeneration, cancellationToken).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<ProfileVersionDto>> ListVersionsAsync(CancellationToken cancellationToken)
    {
        var history = await ReadHistoryAsync(cancellationToken).ConfigureAwait(false);
        return history.OrderByDescending(value => value.CreatedAt).Select(value => new ProfileVersionDto(value.VersionId, value.CreatedAt, value.SourceMessageCount, value.GenerationMethod, value.Notes)).ToArray();
    }

    public async Task<WritingProfileResultDto> RestoreAsync(string versionId, CancellationToken cancellationToken)
    {
        var history = await ReadHistoryAsync(cancellationToken).ConfigureAwait(false);
        var version = history.FirstOrDefault(value => string.Equals(value.VersionId, versionId, StringComparison.Ordinal)) ?? throw Invalid("The requested profile version was not found.");
        var restored = JsonNode.Parse(version.ProfileJson)?.AsObject() ?? throw Invalid("The stored profile version is invalid.");
        Validate(restored);
        await ArchiveCurrentAsync(cancellationToken).ConfigureAwait(false);
        restored["generated_at"] = DateTimeOffset.UtcNow.ToString("O");
        var metadata = restored["generation_metadata"] as JsonObject ?? new JsonObject();
        metadata["restored_from_version"] = versionId;
        metadata["generation_method"] = "restored_profile_version";
        restored["generation_metadata"] = metadata;
        await WriteProfileAsync(restored, cancellationToken).ConfigureAwait(false);
        return await GetAsync(0, cancellationToken).ConfigureAwait(false);
    }

    public JsonObject CreateBaselineProfile(StyleStatisticsDto statistics)
    {
        var greeting = statistics.CommonGreetings.OrderByDescending(value => value.Value).FirstOrDefault().Key;
        var closing = statistics.CommonClosings.OrderByDescending(value => value.Value).FirstOrDefault().Key;
        return new JsonObject
        {
            ["version"] = 1,
            ["generated_at"] = DateTimeOffset.UtcNow.ToString("O"),
            ["source_statistics"] = JsonSerializer.SerializeToNode(statistics),
            ["summary"] = new JsonObject { ["primary_language"] = "Estonian", ["overall_tone"] = "Requires AI-assisted review", ["typical_length"] = statistics.MedianAuthoredCharacters < 500 ? "short" : "medium", ["formality"] = "Requires AI-assisted review" },
            ["style_rules"] = new JsonArray(
                new JsonObject { ["rule"] = "Use the statistically common greeting when context permits.", ["confidence"] = greeting is null ? 0.2 : 0.7, ["evidence_count"] = greeting is null ? 0 : statistics.CommonGreetings[greeting], ["source"] = "automatic_statistics" },
                new JsonObject { ["rule"] = "Use short, readable paragraphs.", ["confidence"] = 0.6, ["evidence_count"] = statistics.WithAuthoredText, ["source"] = "automatic_statistics" }),
            ["common_phrases"] = new JsonArray(),
            ["contextual_variations"] = new JsonArray(),
            ["formatting_preferences"] = new JsonObject { ["most_common_greeting"] = greeting, ["most_common_closing"] = closing },
            ["avoid_patterns"] = new JsonArray(),
            ["user_overrides"] = new JsonArray()
        };
    }

    private JsonObject ParseAndValidate(string json)
    {
        if (string.IsNullOrWhiteSpace(json) || json.Length > 200_000) throw Invalid("profile_json is required and must not exceed 200000 characters.");
        JsonObject profile;
        try { profile = JsonNode.Parse(json)?.AsObject() ?? throw new JsonException(); }
        catch (JsonException ex) { throw new OutlookMcpException(ErrorCodes.InvalidArgument, "profile_json is not valid JSON.", "Provide one structured profile JSON object.", ex); }
        Validate(profile);
        return profile;
    }

    private static void Validate(JsonObject profile)
    {
        foreach (var key in profile.Select(value => value.Key)) if (!AllowedProfileFields.Contains(key)) throw Invalid($"Unsupported profile field: {key}");
        if (profile["version"]?.GetValue<int>() != 1) throw Invalid("Profile version must be 1.");
        if (profile["summary"] is not JsonObject) throw Invalid("Profile summary is required.");
        if (profile["style_rules"] is not JsonArray rules) throw Invalid("Profile style_rules must be an array.");
        var ruleTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var rule in rules.OfType<JsonObject>())
        {
            var text = rule["rule"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(text) || text.Length > 2_000) throw Invalid("Each style rule requires a rule string no longer than 2000 characters.");
            if (!ruleTexts.Add(text)) throw Invalid($"Duplicate style rule: {text}");
            if (rule["confidence"] is JsonValue confidence && (confidence.GetValue<double>() is < 0 or > 1)) throw Invalid("Rule confidence values must be between 0 and 1.");
        }
        if (profile.ToJsonString().Length > 200_000) throw Invalid("The validated profile exceeds 200000 characters.");
    }

    private async Task ArchiveCurrentAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(ProfilePath)) return;
        var json = await File.ReadAllTextAsync(ProfilePath, cancellationToken).ConfigureAwait(false);
        var node = JsonNode.Parse(json)?.AsObject();
        if (node is null) return;
        var metadata = node["generation_metadata"] as JsonObject;
        var history = (await ReadHistoryAsync(cancellationToken).ConfigureAwait(false)).ToList();
        history.Add(new StoredProfileVersion(Guid.NewGuid().ToString("N"), DateTimeOffset.UtcNow, metadata?["source_message_count"]?.GetValue<int>() ?? 0, metadata?["generation_method"]?.GetValue<string>() ?? "unknown", metadata?["notes"]?.GetValue<string>(), json));
        Directory.CreateDirectory(Path.GetDirectoryName(HistoryPath)!);
        await File.WriteAllTextAsync(HistoryPath, JsonSerializer.Serialize(history.TakeLast(100), _jsonOptions), cancellationToken).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<StoredProfileVersion>> ReadHistoryAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(HistoryPath)) return [];
        try { return JsonSerializer.Deserialize<List<StoredProfileVersion>>(await File.ReadAllTextAsync(HistoryPath, cancellationToken).ConfigureAwait(false)) ?? []; }
        catch (JsonException ex) { throw new OutlookMcpException(ErrorCodes.InvalidArgument, "The profile history file is invalid.", "Repair or remove the local profile history file.", ex); }
    }

    private async Task WriteProfileAsync(JsonObject profile, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(ProfilePath)!);
        var temporary = ProfilePath + ".tmp";
        await File.WriteAllTextAsync(temporary, profile.ToJsonString(_jsonOptions), cancellationToken).ConfigureAwait(false);
        File.Move(temporary, ProfilePath, true);
    }

    private static string? Truncate(string? value, int maximum) => value is null || value.Length <= maximum ? value : value[..maximum];
    private static DateTimeOffset? TryDate(string? value) => DateTimeOffset.TryParse(value, out var parsed) ? parsed : null;
    private static OutlookMcpException Invalid(string message) => new(ErrorCodes.InvalidArgument, message, "Correct the structured writing profile and retry.");
    private sealed record StoredProfileVersion(string VersionId, DateTimeOffset CreatedAt, int SourceMessageCount, string GenerationMethod, string? Notes, string ProfileJson);
}
