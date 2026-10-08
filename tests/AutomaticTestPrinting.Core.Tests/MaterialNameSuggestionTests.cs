using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class MaterialNameSuggestionTests
{
    private static readonly string[] TargetAliases = ["英単語ターゲット1900"];
    [Theory]
    [InlineData("古丈上達基礎編45", "古文上達 基礎編")]
    [InlineData("竹岡の英語長文SUREMACY至高の2題", "竹岡の英語長文SUPREMACY至高の20題")]
    public void Suggest_ReturnsCloseNamesWithoutChangingInput(string input, string expected)
    {
        var result = MaterialNameSuggestionService.Suggest(input, [(expected, Array.Empty<string>())]);
        Assert.Equal(expected, Assert.Single(result));
    }

    [Fact]
    public void Suggest_ExactAliasDoesNotNeedCorrection()
    {
        Assert.Empty(MaterialNameSuggestionService.Suggest("英単語ターゲット1900", [("ターゲット1900（6訂版）", TargetAliases)]));
    }

    [Fact]
    public void Suggest_DoesNotSuggestAnotherStageYearOrShortGenericName()
    {
        Assert.Empty(MaterialNameSuggestionService.Suggest("段階突破 英語 2020年度 日大 第2回", [("段階突破 英語 2021年度 日大 第2回", Array.Empty<string>())]));
        Assert.Empty(MaterialNameSuggestionService.Suggest("古文", [("古文上達 基礎編", Array.Empty<string>())]));
    }

    [Fact]
    public void Suggest_UnrelatedNamesAreNotReturned()
    {
        Assert.Empty(MaterialNameSuggestionService.Suggest("未登録教材の名前", [("Vintage 4th Edition", Array.Empty<string>())]));
    }
}
