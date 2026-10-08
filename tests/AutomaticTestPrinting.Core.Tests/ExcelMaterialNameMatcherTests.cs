using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ExcelMaterialNameMatcherTests
{
    [Fact]
    public void ExactSavedReportAliasWinsOverEarlierBuiltInSubstring()
    {
        var builtIn = Profile("builtin", "速読英熟語（新版・改訂版）", "速読英熟語改訂版", "速読英熟語");
        var edited = Profile("custom", "コピー★(改訂版)速読英熟語", "【新版】速読英熟語改訂版");
        foreach (var profiles in new[] { new[] { builtIn, edited }, new[] { edited, builtIn } })
            Assert.Equal(edited, Assert.Single(ExcelMaterialNameMatcher.FindBestMatches(profiles, "【新版】速読英熟語改訂版")));
    }

    [Fact]
    public void MostSpecificSubstringWinsAndAmbiguousExactMatchDoesNotPickFirst()
    {
        var broad = Profile("broad", "別表示", "ターゲット");
        var specific = Profile("specific", "別表示2", "ターゲット1900");
        Assert.Equal(specific, Assert.Single(ExcelMaterialNameMatcher.FindBestMatches([broad, specific], "英単語ターゲット1900 第六版")));
        var duplicate = specific with { Id = "duplicate", DisplayName = "別表示3" };
        Assert.Equal(2, ExcelMaterialNameMatcher.FindBestMatches([specific, duplicate], "ターゲット1900").Count);
    }

    [Fact]
    public void EmptyNamesNeverMatchAndWidthVariantsNormalize()
    {
        var profile = Profile("a", "Ｔａｒｇｅｔ１９００", "");
        Assert.Empty(ExcelMaterialNameMatcher.FindBestMatches([profile], ""));
        Assert.Equal(profile, Assert.Single(ExcelMaterialNameMatcher.FindBestMatches([profile], "Target1900")));
        Assert.Empty(ExcelMaterialNameMatcher.FindBestMatches([profile], "Target1000"));
    }

    private static ExcelTemplateProfile Profile(string id, string name, params string[] aliases) =>
        new() { Id = id, DisplayName = name, MaterialNameKeywords = aliases };
}
