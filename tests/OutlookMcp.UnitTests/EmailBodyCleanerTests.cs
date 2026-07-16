using OutlookMcp.Application.Services;

namespace OutlookMcp.UnitTests;

public sealed class EmailBodyCleanerTests
{
    private readonly EmailBodyCleaner _cleaner = new();

    [Fact]
    public void HtmlToText_RemovesActiveContentAndPreservesStructureAndEstonian()
    {
        const string html = "<html><style>.x{color:red}</style><script>alert(1)</script><body><h1>Tere, Õie</h1><p>Esimene rida<br>teine rida &amp; link</p><ul><li>Üks</li><li>Kaks</li></ul></body></html>";
        var result = _cleaner.Clean(null, html);
        Assert.Contains("Tere, Õie", result.Complete, StringComparison.Ordinal);
        Assert.Contains("teine rida & link", result.Complete, StringComparison.Ordinal);
        Assert.DoesNotContain("alert", result.Complete, StringComparison.Ordinal);
        Assert.DoesNotContain("color:red", result.Complete, StringComparison.Ordinal);
        Assert.Contains('\n', result.Complete);
    }

    [Fact]
    public void Clean_IdentifiesQuotedContentWithoutRemovingIt()
    {
        const string body = "New answer\n\nOn Monday, Client wrote:\n> Keep this quoted line";
        var result = _cleaner.Clean(body, null);
        Assert.True(result.ContainsQuotedContent);
        Assert.Contains("Keep this quoted line", result.Complete, StringComparison.Ordinal);
    }

    [Fact]
    public void Truncate_ReturnsLengthsAndBoundedText()
    {
        var result = EmailBodyCleaner.Truncate("123456789", 5);
        Assert.Equal("12345", result.Value);
        Assert.True(result.Truncated);
        Assert.Equal(9, result.OriginalLength);
    }
}
