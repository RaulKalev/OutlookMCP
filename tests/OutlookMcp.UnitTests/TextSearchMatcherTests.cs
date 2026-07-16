using OutlookMcp.Application.Services;

namespace OutlookMcp.UnitTests;

public sealed class TextSearchMatcherTests
{
    [Fact]
    public void AllTerms_MatchesAcrossDifferentPartsInAnyOrder()
    {
        Assert.True(TextSearchMatcher.Matches("Kaur confirmed the EULE töövestlus", "töövestlus Kaur", "all_terms"));
        Assert.False(TextSearchMatcher.Matches("Kaur confirmed the meeting", "töövestlus Kaur", "all_terms"));
    }

    [Fact]
    public void Phrase_RequiresLiteralSequence()
    {
        Assert.True(TextSearchMatcher.Matches("Project status review", "status review", "phrase"));
        Assert.False(TextSearchMatcher.Matches("Review of project status", "status review", "phrase"));
    }

    [Fact]
    public void EmptyQuery_MatchesWithoutReadingContent()
    {
        Assert.True(TextSearchMatcher.Matches(string.Empty, "  ", "all_terms"));
    }
}
