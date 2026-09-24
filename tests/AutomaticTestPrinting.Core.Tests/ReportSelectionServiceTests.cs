using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReportSelectionServiceTests
{
    [Fact]
    public void CreateDistinct_AcceptsPdfFilesAndRemovesDuplicates()
    {
        var candidates = new[]
        {
            Path.Combine("reports", "student-a.pdf"),
            Path.Combine("reports", "STUDENT-A.PDF"),
            Path.Combine("reports", "notes.txt")
        };

        var result = ReportSelectionService.CreateDistinct(candidates);

        var report = Assert.Single(result);
        Assert.Equal("student-a.pdf", report.FileName, ignoreCase: true);
    }

    [Fact]
    public void CreateDistinct_ExcludesAlreadySelectedReports()
    {
        var path = Path.GetFullPath(Path.Combine("reports", "student-a.pdf"));
        var existing = new[] { new ReportFile(path) };

        var result = ReportSelectionService.CreateDistinct(new[] { path }, existing);

        Assert.Empty(result);
    }
}
