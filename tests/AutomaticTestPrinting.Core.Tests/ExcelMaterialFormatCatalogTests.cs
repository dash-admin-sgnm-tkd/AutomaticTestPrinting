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
    public void CreateOptions_GroupsEquivalentVocabularyProfiles()
    {
        var profiles = LoadBuiltInProfiles();

        var options = ExcelMaterialFormatCatalog.CreateOptions(profiles);

        var vocabulary = Assert.Single(
            options,
            option => option.Id == "vocabulary-standard");
        Assert.Equal("英単語・英熟語（標準型）", vocabulary.DisplayName);
        Assert.Contains("target-1900-sixth-edition", vocabulary.SourceProfileIds);
        Assert.Contains("eiken-pre1-ex-second-edition", vocabulary.SourceProfileIds);
        Assert.Contains("target-1000-fifth-edition", vocabulary.SourceProfileIds);
        Assert.DoesNotContain("vintage-fourth-edition", vocabulary.SourceProfileIds);
    }

    [Fact]
    public void CreateOptions_UsesFriendlyNamesForDistinctLayouts()
    {
        var profiles = LoadBuiltInProfiles();

        var names = ExcelMaterialFormatCatalog.CreateOptions(profiles)
            .Select(option => option.DisplayName)
            .ToArray();

        Assert.Contains("英文法・語法（25問型）", names);
        Assert.Contains("国語・記述（横長解答型）", names);
        Assert.Contains("章番号指定（速読英熟語型）", names);
        Assert.Contains("章・単元指定（公共・政治経済型）", names);
    }

    private static AutomaticTestPrinting.Core.Models.ExcelTemplateProfile[] LoadBuiltInProfiles() =>
        ExcelTemplateProfileStore.LoadDefault().Profiles
            .Where(profile => BuiltInProfileIds.Contains(profile.Id, StringComparer.OrdinalIgnoreCase))
            .ToArray();
}
