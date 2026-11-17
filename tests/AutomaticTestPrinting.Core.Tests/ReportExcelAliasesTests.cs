using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReportExcelAliasesTests
{
    [Fact]
    public void Resolve_UniqueAliasUsesExactWorkbookNameWithOptionalExtension()
    {
        var profile = new ExcelTemplateProfile { Id = "vintage", WorkbookFileName = "★Vintage 4th Edition_20231030本部.xlsm" };
        var result = ReportExcelAliases.Resolve([profile], "Ｖｉｎｔａｇｅ 4th Edition");
        Assert.Same(profile, result.Profile);
        Assert.Null(result.ReviewMessage);
        Assert.Null(ReportExcelAliases.Resolve([profile], "Vintage 4th Edition 派生語").Profile);
    }

    [Fact]
    public void Resolve_ConflictingSheetRowsRequireReviewRatherThanChoosingFirst()
    {
        var result = ReportExcelAliases.Resolve([], "【5訂版】システム英単語");
        Assert.Null(result.Profile);
        Assert.Contains("複数", result.ReviewMessage);
    }

    [Fact]
    public void Resolve_DoesNotTreatDifferentRevisionAsTheSameWorkbook()
    {
        var profile = new ExcelTemplateProfile { Id = "target", WorkbookFileName = "★[5訂版]英熟語ターゲット1000_20220227本部.xlsm" };
        var result = ReportExcelAliases.Resolve([profile], "英熟語ターゲット1000");
        Assert.Null(result.Profile);
        Assert.Contains("202507", result.ReviewMessage);
    }
}
