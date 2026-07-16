using OutlookMcp.Application.Services;

namespace OutlookMcp.UnitTests;

public sealed class SubjectNormalizerTests
{
    [Theory]
    [InlineData("RE: Fwd:  Toila school", "TOILA SCHOOL")]
    [InlineData("VS: SV: Server room", "SERVER ROOM")]
    [InlineData("  Fire   alarm  ", "FIRE ALARM")]
    public void Normalize_RemovesCommonReplyPrefixes(string value, string expected) => Assert.Equal(expected, SubjectNormalizer.Normalize(value));

    [Fact]
    public void ExtractProjectTokens_FindsDistinctCodes()
    {
        var result = SubjectNormalizer.ExtractProjectTokens("Project EULE-2026 and eule-2026, then ABC/123");
        Assert.Equal(["EULE-2026", "ABC/123"], result);
    }
}
