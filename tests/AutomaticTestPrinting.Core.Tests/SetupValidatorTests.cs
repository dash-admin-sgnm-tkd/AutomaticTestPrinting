using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class SetupValidatorTests
{
    [Fact]
    public void Validate_RejectsMissingMaterialFolder()
    {
        var result = SetupValidator.Validate(null, Path.GetTempPath(), []);

        Assert.False(result.IsValid);
        Assert.Contains("教材フォルダー", result.Message);
    }

    [Fact]
    public void Validate_RejectsEmptyReportList()
    {
        var folder = Path.GetTempPath();

        var result = SetupValidator.Validate(folder, folder, []);

        Assert.False(result.IsValid);
        Assert.Contains("レポートPDF", result.Message);
    }

    [Fact]
    public void ValidateFolders_AcceptsAnAvailableFolderWhenAnotherIsMissing()
    {
        var reportPath = Path.GetTempFileName();
        try
        {
            var result = SetupValidator.ValidateFolders(
                [@"Z:\存在しない教材", Path.GetTempPath()],
                Path.GetTempPath(),
                [new AutomaticTestPrinting.Core.Models.ReportFile(reportPath)]);

            Assert.True(result.IsValid);
        }
        finally
        {
            File.Delete(reportPath);
        }
    }
}
