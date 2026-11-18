using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReadingLandscapeCatalogTests
{
    [Theory]
    [InlineData("hijii-reading-essential-extension", 38)]
    [InlineData("hijii-reading-essential-check", 137)]
    [InlineData("hijii-reading-advanced-application", 62)]
    [InlineData("hijii-reading-advanced-check", 62)]
    [InlineData("intro-english-seiko-fourth-five-per-page", 72)]
    [InlineData("basic-lecture-reading-exercise", 72)]
    [InlineData("basic-lecture-reading-examples", 86)]
    [InlineData("polepole-reading-50", 50)]
    [InlineData("interpretation-technique-70", 70)]
    [InlineData("interpretation-technique-100", 100)]
    [InlineData("interpretation-polaris-1-five-per-page", 75)]
    [InlineData("interpretation-polaris-1-original", 75)]
    public void ReadingBoundsIgnoreSerialNumbersOnEmptySourceRows(string id, int lastQuestion)
    {
        var profile = Assert.Single(LoadProfiles(), p => p.Id == id);
        Assert.Equal(lastQuestion, profile.MaximumQuestionNumber);
        Assert.True(profile.UseLandscapeReadingLayout);
        Assert.True(profile.UsesDirectReadingQuestions);
        Assert.False(profile.UsesDirectGrammarQuestions);
        Assert.True(profile.SourceHasSeparateDisplayNumber);
        Assert.Equal(100, profile.EffectiveMaximumQuestionCount);
    }

    [Fact]
    public void RegistrationOffersOneFiveQuestionLandscapeReadingFormat()
    {
        var options = ExcelMaterialFormatCatalog.CreateOptions(LoadProfiles());
        var reading = Assert.Single(options, o => o.Id == "reading-landscape");
        Assert.Equal(12, reading.SourceProfileIds.Count);
        Assert.Contains("1ページ5問", reading.DisplayName);
        Assert.Contains("A4横向き", reading.PreviewDescription);
        Assert.DoesNotContain("hijii-reading-grammar-essential-examples", reading.SourceProfileIds);
    }

    [Theory]
    [InlineData("関正生の英文解釈ポラリス1", "interpretation-polaris-1-five-per-page")]
    [InlineData("ポレポレ", "polepole-reading-50")]
    [InlineData("入門英文問題精講", "intro-english-seiko-fourth-five-per-page")]
    [InlineData("肘井の読解のための英文法必修編確認問題", "hijii-reading-essential-check")]
    public void CanonicalNamesDoNotCollideWithAnotherWorkbook(string name, string expectedId)
    {
        var match = Assert.Single(ExcelMaterialNameMatcher.FindBestMatches(LoadProfiles(), name));
        Assert.Equal(expectedId, match.Id);
    }

    private static IReadOnlyList<AutomaticTestPrinting.Core.Models.ExcelTemplateProfile> LoadProfiles() =>
        ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory, "materials.json")).Profiles;
}
