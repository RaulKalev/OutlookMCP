using Xunit;

namespace OutlookMcp.IntegrationTests;

public sealed class OutlookFactAttribute : FactAttribute
{
    public OutlookFactAttribute(bool writesDraft = false, bool writesFile = false)
    {
        if (!OperatingSystem.IsWindows()) Skip = "Outlook Classic integration tests require Windows.";
        else if (!string.Equals(Environment.GetEnvironmentVariable("OUTLOOK_MCP_INTEGRATION"), "1", StringComparison.Ordinal)) Skip = "Set OUTLOOK_MCP_INTEGRATION=1 to run against a dedicated test profile.";
        else if (writesDraft && !string.Equals(Environment.GetEnvironmentVariable("OUTLOOK_MCP_ALLOW_DRAFT_TESTS"), "1", StringComparison.Ordinal)) Skip = "Set OUTLOOK_MCP_ALLOW_DRAFT_TESTS=1 to allow unsent test drafts.";
        else if (writesFile && !string.Equals(Environment.GetEnvironmentVariable("OUTLOOK_MCP_ALLOW_ATTACHMENT_TESTS"), "1", StringComparison.Ordinal)) Skip = "Set OUTLOOK_MCP_ALLOW_ATTACHMENT_TESTS=1 to allow a test attachment save.";
    }
}
