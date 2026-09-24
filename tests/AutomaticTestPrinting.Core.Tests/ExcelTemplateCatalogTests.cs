using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ExcelTemplateCatalogTests : IDisposable
{
    private readonly string _directory = Path.Combine(
        Path.GetTempPath(),
        $"AutomaticTestPrinting-{Guid.NewGuid():N}");

    [Fact]
    public void Prepare_AcceptsValidTarget1900RequestWhenWorkbookExists()
    {
        Directory.CreateDirectory(_directory);
        var workbookPath = Path.Combine(_directory, "★6訂版)ターゲット1900_本部.xlsm");
        File.WriteAllBytes(workbookPath, []);
        var request = new NormalTestRequestCandidate(
            "英単語ターゲット1900", "501-900", 50, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.True(result.IsSupported);
        Assert.True(result.IsValid);
        Assert.Equal(501, result.StartNumber);
        Assert.Equal(900, result.EndNumber);
        Assert.Equal(workbookPath, result.WorkbookPath);
        Assert.Equal("作業シート", result.Profile?.WorkingSheetName);
        Assert.Equal("G1", result.Profile?.RangeStartCell);
        Assert.Equal("G2", result.Profile?.RangeEndCell);
    }

    [Theory]
    [InlineData("0-100", 10)]
    [InlineData("100-50", 10)]
    [InlineData("1-1901", 10)]
    [InlineData("1-50", 51)]
    [InlineData("1-1900", 101)]
    public void Prepare_RejectsUnsafeRangeOrQuestionCount(string range, int questionCount)
    {
        Directory.CreateDirectory(_directory);
        File.WriteAllBytes(Path.Combine(_directory, "ターゲット1900.xlsm"), []);
        var request = new NormalTestRequestCandidate(
            "英単語ターゲット1900", range, questionCount, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.True(result.IsSupported);
        Assert.False(result.IsValid);
    }

    [Fact]
    public void Prepare_RejectsUnsupportedMaterial()
    {
        var request = new NormalTestRequestCandidate("別の教材", "1-100", 20, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.False(result.IsSupported);
        Assert.False(result.IsValid);
    }

    [Theory]
    [InlineData(
        "英検準1級単熟語EX第2版",
        "コピー(第2版)出る順で最短合格！英検 準1級単熟語EX_202306本部.xlsm",
        "401-1300",
        50,
        "eiken-pre1-ex-second-edition")]
    [InlineData(
        "英熟語ターゲット1000",
        "★【5訂版】英熟語ターゲット1000_202507本部.xlsm",
        "601-1000",
        50,
        "target-1000-fifth-edition")]
    [InlineData(
        "Vintage4thEdition",
        "[新版][4th Edition] Vintage_20231030本部.xlsm",
        "1048-1323",
        25,
        "vintage-fourth-edition")]
    [InlineData(
        "国語力を伸ばす語彙1700(シグマベスト)",
        "国語力を伸ばす 語彙1700_20230301本部.xlsm",
        "1011-1310",
        25,
        "vocabulary-1700-sigma-best")]
    public void Prepare_AcceptsNewlySupportedReportMaterial(
        string materialName,
        string workbookName,
        string range,
        int questionCount,
        string expectedProfileId)
    {
        Directory.CreateDirectory(_directory);
        var workbookPath = Path.Combine(_directory, workbookName);
        File.WriteAllBytes(workbookPath, []);
        var request = new NormalTestRequestCandidate(
            materialName, range, questionCount, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.True(result.IsSupported);
        Assert.True(result.IsValid);
        Assert.Equal(expectedProfileId, result.Profile?.Id);
        Assert.Equal(workbookPath, result.WorkbookPath);
    }

    [Fact]
    public void Prepare_AcceptsVintageNameWithSpacesAndPunctuation()
    {
        Directory.CreateDirectory(_directory);
        File.WriteAllBytes(
            Path.Combine(_directory, "[新版][4th Edition] Vintage_本部.xlsm"), []);
        var request = new NormalTestRequestCandidate(
            "Vintage 4th Edition", "1048-1323", 25, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.True(result.IsSupported);
        Assert.True(result.IsValid);
        Assert.Equal("vintage-fourth-edition", result.Profile?.Id);
    }

    [Fact]
    public void Prepare_RejectsMoreThanTwentyFiveVintageQuestions()
    {
        Directory.CreateDirectory(_directory);
        File.WriteAllBytes(Path.Combine(_directory, "Vintage.xlsm"), []);
        var request = new NormalTestRequestCandidate(
            "Vintage4thEdition", "1-100", 26, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.True(result.IsSupported);
        Assert.False(result.IsValid);
    }

    [Fact]
    public void Prepare_RejectsMoreThanFiftyVocabulary1700Questions()
    {
        Directory.CreateDirectory(_directory);
        File.WriteAllBytes(Path.Combine(_directory, "国語力を伸ばす 語彙1700.xlsm"), []);
        var request = new NormalTestRequestCandidate(
            "国語力を伸ばす語彙1700(シグマベスト)", "1-100", 51, 1, true);

        var result = ExcelTemplateCatalog.Prepare(request, _directory);

        Assert.True(result.IsSupported);
        Assert.False(result.IsValid);
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory))
        {
            Directory.Delete(_directory, true);
        }

        GC.SuppressFinalize(this);
    }
}
