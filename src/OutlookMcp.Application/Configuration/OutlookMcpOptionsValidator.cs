using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Configuration;

public static class OutlookMcpOptionsValidator
{
    public static OutlookMcpOptions Validate(OutlookMcpOptions options)
    {
        if (options.Outlook.DefaultSearchLimit is < 1 or > 100) throw Invalid("Outlook.DefaultSearchLimit must be between 1 and 100.");
        if (options.Outlook.MaximumSearchLimit is < 1 or > 100) throw Invalid("Outlook.MaximumSearchLimit must be between 1 and 100.");
        if (options.Outlook.DefaultSearchLimit > options.Outlook.MaximumSearchLimit) throw Invalid("Outlook.DefaultSearchLimit must not exceed MaximumSearchLimit.");
        if (options.Outlook.DefaultBodyCharacterLimit is < 1 or > 500_000) throw Invalid("Outlook.DefaultBodyCharacterLimit must be between 1 and 500000.");
        if (options.Outlook.OperationTimeoutSeconds is < 1 or > 300) throw Invalid("Outlook.OperationTimeoutSeconds must be between 1 and 300.");
        if (options.Outlook.MaximumRecursiveFolders is < 1 or > 10_000) throw Invalid("Outlook.MaximumRecursiveFolders must be between 1 and 10000.");
        if (options.Outlook.MaximumBatchSize is < 1 or > 500) throw Invalid("Outlook.MaximumBatchSize must be between 1 and 500.");
        if (options.Outlook.AllowAttachmentSaving && options.Outlook.AllowedAttachmentDirectories.Count == 0) throw Invalid("At least one allowed attachment directory is required when attachment saving is enabled.");
        if (options.WritingStyle.BatchSize is < 1 or > 500) throw Invalid("WritingStyle.BatchSize must be between 1 and 500.");
        if (string.IsNullOrWhiteSpace(options.WritingStyle.DatabasePath) || string.IsNullOrWhiteSpace(options.WritingStyle.ProfilePath) || string.IsNullOrWhiteSpace(options.WritingStyle.ProfileHistoryPath)) throw Invalid("WritingStyle database and profile paths are required.");
        if (options.WritingStyle.DelayBetweenBatchesMilliseconds is < 0 or > 10_000) throw Invalid("WritingStyle.DelayBetweenBatchesMilliseconds must be between 0 and 10000.");
        if (options.WritingStyle.MaximumRetrievedExamples is < 1 or > 10) throw Invalid("WritingStyle.MaximumRetrievedExamples must be between 1 and 10.");
        if (options.WritingStyle.MaximumCharactersPerExample is < 100 or > 20_000) throw Invalid("WritingStyle.MaximumCharactersPerExample must be between 100 and 20000.");
        if (options.WritingStyle.MaximumDraftContextCharacters is < 2_000 or > 100_000) throw Invalid("WritingStyle.MaximumDraftContextCharacters must be between 2000 and 100000.");
        if (options.WritingStyle.ProfileRefreshRecommendationThreshold is < 1 or > 100_000) throw Invalid("WritingStyle.ProfileRefreshRecommendationThreshold must be between 1 and 100000.");
        if (options.WritingStyle.RecencyWeighting.RecentMonths is < 1 or > 120) throw Invalid("WritingStyle.RecencyWeighting.RecentMonths must be between 1 and 120.");
        if (options.WritingStyle.RecencyWeighting.MiddleYears is < 1 or > 50) throw Invalid("WritingStyle.RecencyWeighting.MiddleYears must be between 1 and 50.");
        if (options.WritingStyle.AdditionalSignatureMarkers.Concat(options.WritingStyle.AdditionalDisclaimerMarkers).Any(value => string.IsNullOrWhiteSpace(value) || value.Length > 500)) throw Invalid("WritingStyle signature/disclaimer markers must contain 1-500 characters.");
        ValidateWeight(options.WritingStyle.RecencyWeighting.RecentWeight, "RecentWeight");
        ValidateWeight(options.WritingStyle.RecencyWeighting.MiddleWeight, "MiddleWeight");
        ValidateWeight(options.WritingStyle.RecencyWeighting.OlderWeight, "OlderWeight");
        if (options.Logging.RetentionDays is < 1 or > 365) throw Invalid("Logging.RetentionDays must be between 1 and 365.");
        return options;
    }

    private static OutlookMcpException Invalid(string message) => new(ErrorCodes.InvalidArgument, message, "Correct the per-user config.json file and retry.");
    private static void ValidateWeight(double value, string name)
    {
        if (value is < 0.1 or > 2.0) throw Invalid($"WritingStyle.RecencyWeighting.{name} must be between 0.1 and 2.0.");
    }
}
