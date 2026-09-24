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

    public void Dispose()
    {
        if (Directory.Exists(_directory))
        {
            Directory.Delete(_directory, true);
        }

        GC.SuppressFinalize(this);
    }
}
