using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.UnitTests;

public sealed class FolderRuleMatcherTests
{
    [Fact]
    public void Matches_UsesOrWithinAConditionAndAndAcrossConditions()
    {
        var rule = new CreateFolderRuleRequest("store", "folder", "Invoices",
            SenderAddressContains: ["billing@example.com", "accounts@example.com"],
            SubjectContains: ["invoice", "receipt"]);

        Assert.True(FolderRuleMatcher.Matches(rule, "billing@example.com", "Your receipt", null));
        Assert.False(FolderRuleMatcher.Matches(rule, "news@example.com", "Your receipt", null));
        Assert.False(FolderRuleMatcher.Matches(rule, "billing@example.com", "Welcome", null));
    }

    [Fact]
    public void Matches_BodyOrSubjectChecksEitherFieldIgnoringCase()
    {
        var rule = new CreateFolderRuleRequest("store", "folder", "Project",
            BodyOrSubjectContains: ["ABC-123"]);

        Assert.True(FolderRuleMatcher.Matches(rule, null, "Update for abc-123", null));
        Assert.True(FolderRuleMatcher.Matches(rule, null, "Update", "Work on ABC-123 is complete."));
        Assert.False(FolderRuleMatcher.Matches(rule, null, "Update", "No project code here."));
    }

    [Fact]
    public void Matches_UsesSenderSubstringSemanticsLikeOutlook()
    {
        var rule = new CreateFolderRuleRequest("store", "folder", "Domain",
            SenderAddressContains: ["@example.com"]);

        Assert.True(FolderRuleMatcher.Matches(rule, "person@example.com", null, null));
        Assert.False(FolderRuleMatcher.Matches(rule, "person@example.org", null, null));
    }
}
