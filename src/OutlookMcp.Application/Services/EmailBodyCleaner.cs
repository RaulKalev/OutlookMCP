using System.Net;
using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;

namespace OutlookMcp.Application.Services;

public sealed record CleanedBody(string Complete, string Preview, bool ContainsQuotedContent, bool ContainsLikelySignature);

public sealed partial class EmailBodyCleaner
{
    private readonly HtmlParser _parser = new();

    [GeneratedRegex(@"\r?\n[ \t]*\r?\n(?:[ \t]*\r?\n)+", RegexOptions.CultureInvariant)]
    private static partial Regex ExcessBlankLines();

    [GeneratedRegex(@"(?im)^(>+\s|On .+wrote:|From:\s.+|-----Original Message-----|Saatja:\s.+)", RegexOptions.CultureInvariant)]
    private static partial Regex QuoteMarker();

    [GeneratedRegex(@"(?im)^\s*(--\s*$|Regards[,!]?\s*$|Best regards[,!]?\s*$|Lugupidamisega[,!]?\s*$)", RegexOptions.CultureInvariant)]
    private static partial Regex SignatureMarker();

    public CleanedBody Clean(string? plainText, string? html, int previewLength = 500)
    {
        var text = !string.IsNullOrWhiteSpace(plainText) ? plainText! : HtmlToText(html ?? string.Empty);
        text = text.Replace("\0", string.Empty, StringComparison.Ordinal).Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        text = ExcessBlankLines().Replace(text, "\n\n").Trim();
        var preview = text.Length <= previewLength ? text : text[..previewLength].TrimEnd() + "…";
        return new(text, preview, QuoteMarker().IsMatch(text), SignatureMarker().IsMatch(text));
    }

    public string HtmlToText(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;
        var document = _parser.ParseDocument(html);
        foreach (var element in document.QuerySelectorAll("script,style,noscript,svg,canvas,iframe,object,embed,meta,link")) element.Remove();
        foreach (var br in document.QuerySelectorAll("br")) br.Replace(document.CreateTextNode("\n"));
        foreach (var block in document.QuerySelectorAll("p,div,li,tr,h1,h2,h3,h4,h5,h6,blockquote")) block.Append(document.CreateTextNode("\n"));
        return WebUtility.HtmlDecode(document.Body?.TextContent ?? document.DocumentElement.TextContent);
    }

    public static (string Value, bool Truncated, int OriginalLength) Truncate(string? value, int maximum)
    {
        var text = value ?? string.Empty;
        ArgumentOutOfRangeException.ThrowIfLessThan(maximum, 1);
        return text.Length <= maximum ? (text, false, text.Length) : (text[..maximum], true, text.Length);
    }
}
