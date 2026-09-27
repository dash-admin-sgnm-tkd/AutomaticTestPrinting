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

        var vocabulary = Assert.Single(
            options,
            option => option.Id == "word-pair-list");
        Assert.Equal("英単語・英熟語（1問1答）", vocabulary.DisplayName);
        Assert.Contains("target-1900-sixth-edition", vocabulary.SourceProfileIds);
        Assert.Contains("eiken-pre1-ex-second-edition", vocabulary.SourceProfileIds);
        Assert.Contains("target-1000-fifth-edition", vocabulary.SourceProfileIds);
        Assert.DoesNotContain("vintage-fourth-edition", vocabulary.SourceProfileIds);

        var grammar = Assert.Single(
            options,
            option => option.Id == "grammar-choice");
        Assert.Equal("文法・選択問題（問題文ワイド）", grammar.DisplayName);
        Assert.Contains("vintage-fourth-edition", grammar.SourceProfileIds);
    }

    [Fact]
    public void CreateOptions_UsesFriendlyNamesForDistinctLayouts()
    {
        var profiles = LoadBuiltInProfiles();

        var names = ExcelMaterialFormatCatalog.CreateOptions(profiles)
            .Select(option => option.DisplayName)
            .ToArray();

        Assert.Contains("英単語・英熟語（1問1答）", names);
        Assert.Contains("文法・選択問題（問題文ワイド）", names);
        Assert.Contains("国語（語彙・漢字・古文の1問1答）", names);
        Assert.Contains("問題・解答（均等2列）", names);
        Assert.Contains("長文問題（問題欄ワイド）", names);
    }

    private static AutomaticTestPrinting.Core.Models.ExcelTemplateProfile[] LoadBuiltInProfiles() =>
        ExcelTemplateProfileStore.LoadDefault().Profiles
            .Where(profile => BuiltInProfileIds.Contains(profile.Id, StringComparer.OrdinalIgnoreCase))
            .ToArray();
}
