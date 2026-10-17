using System.Reflection;
using ModelContextProtocol.Server;
using OutlookMcp.Server.Mcp;

namespace OutlookMcp.UnitTests;

public sealed class McpToolSurfaceTests
{
    private static readonly string[] MailboxChangingVerbs = ["create", "move", "save", "send", "delete", "forward", "login", "logout"];

    [Fact]
    public void AssistantProfile_HasSeventeenReadOnlyTools()
    {
        var names = ToolNames(typeof(OutlookAssistantTools));

        Assert.Equal(17, names.Count);
        Assert.Contains("outlook_read_attachment_text", names);
        Assert.DoesNotContain(names, name => MailboxChangingVerbs.Any(verb => name.Split('_').Contains(verb)));
    }

    [Fact]
    public void AttachmentTextTool_IsInFullAndMailButNotCompact()
    {
        Assert.Contains("outlook_read_attachment_text", ToolNames(typeof(OutlookTools)));
        Assert.DoesNotContain("outlook_read_attachment_text", ToolNames(typeof(OutlookCompactTools)));
    }

    private static List<string> ToolNames(Type type) => type
        .GetMethods(BindingFlags.Public | BindingFlags.Instance)
        .Select(method => method.GetCustomAttribute<McpServerToolAttribute>()?.Name)
        .OfType<string>()
        .ToList();
}
