using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;

namespace OutlookMcp.UnitTests;

public sealed class ErrorAndConfigurationTests
{
    [Fact]
    public void ErrorConversion_HidesTechnicalDetailsByDefault()
    {
        var exception = new OutlookMcpException("CODE", "Safe message", "Retry", new InvalidOperationException("secret technical text"));
        var safe = exception.ToError(false);
        var verbose = exception.ToError(true);
        Assert.Null(safe.TechnicalDetails);
        Assert.Equal("secret technical text", verbose.TechnicalDetails);
    }

    [Fact]
    public void Configuration_DefaultsAreSafetyBounded()
    {
        var options = new OutlookMcpOptions();
        Assert.Equal(100, options.Outlook.MaximumSearchLimit);
        Assert.False(options.Outlook.AllowHtmlBody);
        Assert.True(options.Outlook.AllowAttachmentSaving);
        Assert.False(options.Logging.IncludeTechnicalDetails);
        Assert.NotEmpty(options.Outlook.AllowedAttachmentDirectories);
    }

    [Fact]
    public void Configuration_RejectsInvalidLimits()
    {
        var options = new OutlookMcpOptions();
        options.Outlook.MaximumSearchLimit = 101;
        Assert.Throws<OutlookMcpException>(() => OutlookMcpOptionsValidator.Validate(options));
    }
}
