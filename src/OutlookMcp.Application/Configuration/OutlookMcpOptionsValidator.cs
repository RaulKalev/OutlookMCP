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
        if (options.Logging.RetentionDays is < 1 or > 365) throw Invalid("Logging.RetentionDays must be between 1 and 365.");
        return options;
    }

    private static OutlookMcpException Invalid(string message) => new(ErrorCodes.InvalidArgument, message, "Correct the per-user config.json file and retry.");
}
