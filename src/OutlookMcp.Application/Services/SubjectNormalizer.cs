using System.Text.RegularExpressions;

namespace OutlookMcp.Application.Services;

public static partial class SubjectNormalizer
{
    [GeneratedRegex(@"^\s*((re|fw|fwd|vs|sv|aw)\s*:\s*)+", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex PrefixRegex();

    [GeneratedRegex(@"\s+", RegexOptions.CultureInvariant)]
    private static partial Regex SpaceRegex();

    public static string Normalize(string? subject) =>
        SpaceRegex().Replace(PrefixRegex().Replace(subject ?? string.Empty, string.Empty), " ").Trim().ToUpperInvariant();

    public static IReadOnlyList<string> ExtractProjectTokens(string? text) =>
        Regex.Matches(text ?? string.Empty, @"\b[A-ZÄÖÜÕ]{2,10}[-_/]\d{2,8}\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
            .Select(match => match.Value.ToUpperInvariant()).Distinct(StringComparer.Ordinal).ToArray();
}
