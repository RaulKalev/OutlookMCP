using OutlookMcp.Application.Errors;
using OutlookMcp.Server.Mcp;

namespace OutlookMcp.UnitTests;

public sealed class McpToolProfileParserTests
{
    [Fact]
    public void Parse_AssistantProfile()
    {
        Assert.Equal(McpToolProfile.Assistant, McpToolProfileParser.Parse(["--tool-profile", "assistant"]));
    }

    [Fact]
    public void Parse_DefaultsToFull()
    {
        Assert.Equal(McpToolProfile.Full, McpToolProfileParser.Parse([]));
    }

    [Fact]
    public void Parse_RejectsUnknownProfile()
    {
        Assert.Throws<OutlookMcpException>(() => McpToolProfileParser.Parse(["--tool-profile", "writer"]));
    }
}
