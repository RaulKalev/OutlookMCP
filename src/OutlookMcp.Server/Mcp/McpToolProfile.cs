using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Server.Mcp;

internal enum McpToolProfile
{
    Compact,
    Mail,
    Style,
    Full
}

internal static class McpToolProfileParser
{
    public static McpToolProfile Parse(string[] args)
    {
        var index = Array.FindIndex(args, value =>
            string.Equals(value, "--tool-profile", StringComparison.OrdinalIgnoreCase));
        if (index < 0) return McpToolProfile.Full;
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
        {
            throw Invalid();
        }

        return args[index + 1].ToLowerInvariant() switch
        {
            "compact" => McpToolProfile.Compact,
            "mail" => McpToolProfile.Mail,
            "style" => McpToolProfile.Style,
            "full" => McpToolProfile.Full,
            _ => throw Invalid()
        };
    }

    private static OutlookMcpException Invalid() => new(
        ErrorCodes.InvalidArgument,
        "--tool-profile must be compact, mail, style, or full.",
        "Choose a supported MCP tool profile and restart the server.");
}
