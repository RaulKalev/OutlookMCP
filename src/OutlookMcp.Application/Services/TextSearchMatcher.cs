namespace OutlookMcp.Application.Services;

public static class TextSearchMatcher
{
    public static bool Matches(string source, string? query, string mode)
    {
        if (string.IsNullOrWhiteSpace(query)) return true;
        if (mode == "phrase") return source.Contains(query.Trim(), StringComparison.CurrentCultureIgnoreCase);

        return query.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .All(term => source.Contains(term, StringComparison.CurrentCultureIgnoreCase));
    }
}
