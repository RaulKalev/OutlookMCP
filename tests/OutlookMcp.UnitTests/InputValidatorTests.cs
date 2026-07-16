using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.UnitTests;

public sealed class InputValidatorTests
{
    private readonly OutlookOptions _options = new() { MaximumSearchLimit = 100 };

    [Fact]
    public void Validate_AcceptsBoundedSearch()
    {
        var request = new SearchEmailsRequest("Toila", DateFrom: DateTimeOffset.UtcNow.AddDays(-1), DateTo: DateTimeOffset.UtcNow, MaxResults: 25);
        Assert.Same(request, InputValidator.Validate(request, _options));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(101)]
    public void Validate_RejectsOutOfRangeLimits(int maximum)
    {
        var exception = Assert.Throws<OutlookMcpException>(() => InputValidator.Validate(new SearchEmailsRequest("x", MaxResults: maximum), _options));
        Assert.Equal(ErrorCodes.ResultLimitExceeded, exception.Code);
    }

    [Fact]
    public void Validate_RejectsReverseDateRange()
    {
        var request = new SearchEmailsRequest("x", DateFrom: DateTimeOffset.UtcNow, DateTo: DateTimeOffset.UtcNow.AddDays(-1));
        var exception = Assert.Throws<OutlookMcpException>(() => InputValidator.Validate(request, _options));
        Assert.Equal(ErrorCodes.InvalidArgument, exception.Code);
    }

    [Fact]
    public void ValidateRecipients_AcceptsEstonianDisplayName()
    {
        InputValidator.ValidateRecipients("Märt Tamm <mart@example.ee>; liis@example.ee");
    }

    [Fact]
    public void ValidateRecipients_RejectsInvalidAddress()
    {
        Assert.Throws<OutlookMcpException>(() => InputValidator.ValidateRecipients("not an address"));
    }

    [Fact]
    public void ValidateSearch_RejectsUnknownQueryMode()
    {
        var request = new SearchEmailsRequest("test", QueryMode: "fuzzy");
        Assert.Throws<OutlookMcpException>(() => InputValidator.Validate(request, _options));
    }

    [Fact]
    public void ValidateBatch_RejectsDuplicatesAndOversizedRequests()
    {
        Assert.Throws<OutlookMcpException>(() => InputValidator.ValidateBatch(["one", "one"], 10));
        Assert.Throws<OutlookMcpException>(() => InputValidator.ValidateBatch(["one", "two"], 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("Nested\\Folder")]
    [InlineData(" trailing ")]
    public void ValidateFolderName_RejectsAmbiguousNames(string value)
    {
        Assert.Throws<OutlookMcpException>(() => InputValidator.ValidateFolderName(value));
    }
}
