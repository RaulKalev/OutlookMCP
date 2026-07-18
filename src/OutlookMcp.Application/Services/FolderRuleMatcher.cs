using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

public static class FolderRuleMatcher
{
    public static bool Matches(CreateFolderRuleRequest rule, string? senderAddress, string? subject, string? body)
    {
        return MatchesAny(senderAddress, rule.SenderAddressContains)
            && MatchesAny(subject, rule.SubjectContains)
            && MatchesAny(body, rule.BodyContains)
            && MatchesBodyOrSubject(subject, body, rule.BodyOrSubjectContains);
    }

    private static bool MatchesAny(string? candidate, IReadOnlyList<string>? values)
    {
        return values is null || values.Count == 0 || values.Any(value => Contains(candidate, value));
    }

    private static bool MatchesBodyOrSubject(string? subject, string? body, IReadOnlyList<string>? values)
    {
        return values is null || values.Count == 0 || values.Any(value => Contains(subject, value) || Contains(body, value));
    }

    private static bool Contains(string? candidate, string value) =>
        candidate?.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
}
