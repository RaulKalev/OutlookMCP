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
        Assert.Equal(100, options.Outlook.MaximumBatchSize);
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

    [Fact]
    public void Configuration_CalendarSyncDefaultsAreValid()
    {
        var options = new OutlookMcpOptions();
        Assert.Equal(3, options.CalendarSync.DefaultMonthsAhead);
        Assert.Null(options.CalendarSync.SourceCalendarFolderId);
        Assert.Null(options.CalendarSync.ClientId);
        Assert.Equal("common", options.CalendarSync.TenantId);
        OutlookMcpOptionsValidator.Validate(options);
    }

    [Fact]
    public void Configuration_RejectsEmptyCalendarSyncTenantAndCachePath()
    {
        var options = new OutlookMcpOptions();
        options.CalendarSync.TenantId = " ";
        Assert.Throws<OutlookMcpException>(() => OutlookMcpOptionsValidator.Validate(options));
        options.CalendarSync.TenantId = "common";
        options.CalendarSync.TokenCacheDirectory = "";
        Assert.Throws<OutlookMcpException>(() => OutlookMcpOptionsValidator.Validate(options));
    }

    [Theory]
    [InlineData(0, 24, 2_500)]
    [InlineData(25, 24, 2_500)]
    [InlineData(3, 37, 2_500)]
    [InlineData(3, 24, 99)]
    [InlineData(3, 24, 20_001)]
    public void Configuration_RejectsInvalidCalendarSyncLimits(int defaultMonths, int maximumMonths, int maximumItems)
    {
        var options = new OutlookMcpOptions();
        options.CalendarSync.DefaultMonthsAhead = defaultMonths;
        options.CalendarSync.MaximumMonthsAhead = maximumMonths;
        options.CalendarSync.MaximumItemsScanned = maximumItems;
        Assert.Throws<OutlookMcpException>(() => OutlookMcpOptionsValidator.Validate(options));
    }
}
