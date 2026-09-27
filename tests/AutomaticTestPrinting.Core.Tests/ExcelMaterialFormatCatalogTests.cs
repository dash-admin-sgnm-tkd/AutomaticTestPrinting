using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ExcelMaterialFormatCatalogTests
{
    private static readonly string[] BuiltInProfileIds =
    [
        "target-1900-sixth-edition",
        "eiken-pre1-ex-second-edition",
        "target-1000-fifth-edition",
        "vintage-fourth-edition",
        "vocabulary-1700-sigma-best",
        "rapid-reading-idioms-revised",
        "civic-politics-economics-fifth-edition"
    ];

    [Fact]
    public void CreateOptions_GroupsProfilesWithTheSameFinishedLayout()
    {
        var profiles = LoadBuiltInProfiles();

        var options = ExcelMaterialFormatCatalog.CreateOptions(profiles);

        var standard = Assert.Single(
            options,
            option => option.Id == "standard-table");
        Assert.Equal("標準（問題・解答の表）", standard.DisplayName);
        Assert.Contains("target-1900-sixth-edition", standard.SourceProfileIds);
        Assert.Contains("eiken-pre1-ex-second-edition", standard.SourceProfileIds);
        Assert.Contains("target-1000-fifth-edition", standard.SourceProfileIds);
        Assert.Contains("vintage-fourth-edition", standard.SourceProfileIds);
    }

    [Fact]
    public void CreateOptions_UsesFriendlyNamesForDistinctLayouts()
    {
        var profiles = LoadBuiltInProfiles();

        var names = ExcelMaterialFormatCatalog.CreateOptions(profiles)
            .Select(option => option.DisplayName)
            .ToArray();

        Assert.Contains("標準（問題・解答の表）", names);
        Assert.Contains("解答欄ワイド（解答を広く表示）", names);
        Assert.Contains("均等2列（問題・解答が同じ幅）", names);
        Assert.Contains("問題欄ワイド（長文・自動行高）", names);
    }

    private static AutomaticTestPrinting.Core.Models.ExcelTemplateProfile[] LoadBuiltInProfiles() =>
        ExcelTemplateProfileStore.LoadDefault().Profiles
            .Where(profile => BuiltInProfileIds.Contains(profile.Id, StringComparer.OrdinalIgnoreCase))
            .ToArray();
}
