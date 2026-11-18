using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReadingPortraitCatalogTests
{
    [Theory]
    [InlineData("interpretation-polaris-2", 55)]
    [InlineData("intro-english-seiko-fourth-portrait", 72)]
    [InlineData("hitotsu-high-school-reading-basic", 236)]
    [InlineData("hitotsu-high-school-reading-review", 51)]
    public void PortraitProfilesUseActualQuestionBoundsAndTenRows(string id, int end)
    {
        var profiles = ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory, "materials.json")).Profiles;
        var profile = Assert.Single(profiles, p => p.Id == id);
        Assert.Equal(end, profile.MaximumQuestionNumber);
        Assert.True(profile.UsesDirectReadingQuestions);
        Assert.True(profile.UsePortraitReadingLayout);
        Assert.False(profile.UseLandscapeReadingLayout);
        Assert.Equal(10, profile.OriginalGrammarQuestionsPerPage);
        Assert.Equal(130, profile.OriginalGrammarRowHeight);
        var option = Assert.Single(ExcelMaterialFormatCatalog.CreateOptions(profiles), o => o.Id == "reading-portrait");
        Assert.Equal(4, option.SourceProfileIds.Count);
        Assert.Contains("1ページ10問", option.DisplayName);
    }

    [Fact]
    public void ReadingFormattingRemovesOnlyBlankLinesAndPresentationPadding()
    {
        const string text = "日本語の指示\r\n\r\nWhether the resources are materials or energy,\r\n\r\ntechnology is needed.　　　　　　　　　*resources「資源」　\r\n";
        var result = ReadingTextFormatter.Normalize(text, reduceInlinePadding: true);
        Assert.Equal("日本語の指示\nWhether the resources are materials or energy,\ntechnology is needed. *resources「資源」", result);
    }

    [Fact]
    public void AnswerFormattingKeepsSentenceAnalysisAlignmentAndSingleNewlines()
    {
        const string text = "英文\n   S       V    C\n\n日本語訳\n";
        Assert.Equal("英文\n   S       V    C\n日本語訳", ReadingTextFormatter.Normalize(text, reduceInlinePadding: false));
    }
}
