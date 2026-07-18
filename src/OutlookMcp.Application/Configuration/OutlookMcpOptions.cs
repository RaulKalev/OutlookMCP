namespace OutlookMcp.Application.Configuration;

public sealed class OutlookMcpOptions
{
    public OutlookOptions Outlook { get; set; } = new();
    public CalendarSyncOptions CalendarSync { get; set; } = new();
    public WritingStyleOptions WritingStyle { get; set; } = new();
    public LoggingOptions Logging { get; set; } = new();
}

public sealed class OutlookOptions
{
    public bool StartIfNotRunning { get; set; } = true;
    public List<string> AllowedStores { get; set; } = [];
    public List<string> BlockedFolderPaths { get; set; } = ["\\Personal"];
    public List<string> DefaultSearchFolders { get; set; } = ["Inbox", "Sent Items"];
    public int DefaultSearchLimit { get; set; } = 25;
    public int MaximumSearchLimit { get; set; } = 100;
    public int DefaultBodyCharacterLimit { get; set; } = 50_000;
    public int OperationTimeoutSeconds { get; set; } = 30;
    public bool AllowHtmlBody { get; set; }
    public bool AllowAttachmentSaving { get; set; } = true;
    public List<string> AllowedAttachmentDirectories { get; set; } = ["%USERPROFILE%\\Downloads\\Outlook AI"];
    public bool AllowSelectedEmailAccess { get; set; } = true;
    public int MaximumRecursiveFolders { get; set; } = 1_000;
    public int MaximumBatchSize { get; set; } = 100;
}

public sealed class CalendarSyncOptions
{
    public string? ClientId { get; set; }
    public string TenantId { get; set; } = "common";
    public string? SourceCalendarFolderId { get; set; }
    public string? SourceStoreId { get; set; }
    public string? TargetCalendarId { get; set; }
    public int DefaultMonthsAhead { get; set; } = 3;
    public int MaximumMonthsAhead { get; set; } = 24;
    public int MaximumItemsScanned { get; set; } = 2_500;
    public string TokenCacheDirectory { get; set; } = "%APPDATA%\\EULE\\OutlookMcp\\Exchange";
}

public sealed class LoggingOptions
{
    public string Level { get; set; } = "Information";
    public int RetentionDays { get; set; } = 14;
    public bool IncludeTechnicalDetails { get; set; }
}

public sealed class WritingStyleOptions
{
    public bool Enabled { get; set; } = true;
    public string DatabasePath { get; set; } = "%APPDATA%\\EULE\\OutlookMcp\\WritingStyle\\sent-email-index.db";
    public string ProfilePath { get; set; } = "%APPDATA%\\EULE\\OutlookMcp\\WritingStyle\\writing-profile.json";
    public string ProfileHistoryPath { get; set; } = "%APPDATA%\\EULE\\OutlookMcp\\WritingStyle\\writing-profile-history.json";
    public bool ScanAllSentFolders { get; set; } = true;
    public List<string> AllowedStores { get; set; } = [];
    public int BatchSize { get; set; } = 100;
    public int DelayBetweenBatchesMilliseconds { get; set; } = 100;
    public bool SaveCheckpointAfterEachBatch { get; set; } = true;
    public bool SyncBeforeDraftContext { get; set; } = true;
    public int MaximumRetrievedExamples { get; set; } = 5;
    public int MaximumCharactersPerExample { get; set; } = 3_000;
    public int MaximumDraftContextCharacters { get; set; } = 30_000;
    public bool AllowFullAuthoredTextInResponses { get; set; }
    public bool StoreCleanBody { get; set; } = true;
    public bool StoreQuotedText { get; set; } = true;
    public bool StoreSignatureText { get; set; } = true;
    public List<string> AdditionalSignatureMarkers { get; set; } = [];
    public List<string> AdditionalDisclaimerMarkers { get; set; } = [];
    public bool EnableFullTextSearch { get; set; } = true;
    public int ProfileRefreshRecommendationThreshold { get; set; } = 100;
    public RecencyWeightingOptions RecencyWeighting { get; set; } = new();
}

public sealed class RecencyWeightingOptions
{
    public int RecentMonths { get; set; } = 12;
    public double RecentWeight { get; set; } = 1.0;
    public int MiddleYears { get; set; } = 3;
    public double MiddleWeight { get; set; } = 0.8;
    public double OlderWeight { get; set; } = 0.6;
}

public static class AppPaths
{
    public static string ConfigDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "EULE", "OutlookMcp");
    public static string ConfigPath => Path.Combine(ConfigDirectory, "config.json");
    public static string LogDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EULE", "OutlookMcp", "Logs");
    public static string ExpandPath(string value) => Path.GetFullPath(Environment.ExpandEnvironmentVariables(value));
}
