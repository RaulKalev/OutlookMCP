using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;

namespace OutlookMcp.UnitTests;

public sealed class AttachmentPathPolicyTests
{
    [Fact]
    public void SanitizeFileName_RemovesTraversalAndInvalidCharacters()
    {
        var result = AttachmentPathPolicy.SanitizeFileName("..\\..\\client:report?.pdf");
        Assert.DoesNotContain("..", result, StringComparison.Ordinal);
        Assert.DoesNotContain('\\', result);
        Assert.DoesNotContain(':', result);
        Assert.EndsWith(".pdf", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateDirectory_AllowsConfiguredDirectoryAndChild()
    {
        var root = Path.Combine(Path.GetTempPath(), "outlook-mcp-tests", Guid.NewGuid().ToString("N"));
        var policy = new AttachmentPathPolicy(new OutlookOptions { AllowedAttachmentDirectories = [root] });
        Assert.Equal(Path.Combine(root, "child"), policy.ValidateDirectory(Path.Combine(root, "child")));
    }

    [Fact]
    public void ValidateDirectory_RejectsSiblingWithCommonPrefix()
    {
        var root = Path.Combine(Path.GetTempPath(), "outlook-mcp-tests", "allowed");
        var policy = new AttachmentPathPolicy(new OutlookOptions { AllowedAttachmentDirectories = [root] });
        var exception = Assert.Throws<OutlookMcpException>(() => policy.ValidateDirectory(root + "-evil"));
        Assert.Equal("ATTACHMENT_PATH_NOT_ALLOWED", exception.Code);
    }

    [Fact]
    public void SanitizeFileName_PrefixesReservedWindowsDeviceNames()
    {
        Assert.Equal("_CON.txt", AttachmentPathPolicy.SanitizeFileName("CON.txt"));
    }
}
