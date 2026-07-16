namespace OutlookMcp.Application.Configuration;

public sealed class OutlookMcpOptions
{
    public OutlookOptions Outlook { get; set; } = new();
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
}

public sealed class LoggingOptions
{
    public string Level { get; set; } = "Information";
    public int RetentionDays { get; set; } = 14;
    public bool IncludeTechnicalDetails { get; set; }
}

public static class AppPaths
{
    public static string ConfigDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "EULE", "OutlookMcp");
    public static string ConfigPath => Path.Combine(ConfigDirectory, "config.json");
    public static string LogDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EULE", "OutlookMcp", "Logs");
}
