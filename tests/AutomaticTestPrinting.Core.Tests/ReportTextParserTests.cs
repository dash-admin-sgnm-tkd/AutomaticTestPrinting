using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReportTextParserTests
{
    [Fact]
    public void Parse_HandlesSpacesInsertedByJapaneseOcr()
    {
        var pages = new[]
        {
            new RecognizedReportPage(
                1,
                180,
                "生 徒 名 : 相 川 創 太 講 師 名 : 嶺 脇 結 花 前 回 か ら の 宿 題 テ ス ト 作 成 依 頼 100 問"),
            new RecognizedReportPage(
                2,
                180,
                "次 回 ま で の 宿 題 テ ス ト 作 成 依 頼 50 問 25 問 25 問 段 階 突 破")
        };

        var result = ReportTextParser.Parse("report.pdf", pages);

        Assert.Equal("相川創太", result.StudentName);
        Assert.Equal("問題数のOCR候補：50問・25問・25問（原本確認）", result.NormalTestSummary);
        Assert.Equal("段階突破の記載あり（内容を確認してください）", result.StageTestSummary);
    }

    [Fact]
    public void Parse_UsesTheLastStageTestDetailsAfterNextContent()
    {
        var pages = new[]
        {
            new RecognizedReportPage(
                1,
                0,
                "生徒名. 石村 光彩 2026年09月29日 講師名: 矢田 遥菜 次回までの宿題 テスト作成依頼"),
            new RecognizedReportPage(
                2,
                0,
                "段階突破テスト 科目 国語（古文） 次回の内容 科目 国語（古文） 得点77% 年度2021 段階 日大 回1 年度2020 段階 日大 回2")
        };

        var result = ReportTextParser.Parse("report.pdf", pages);

        Assert.Equal("石村光彩", result.StudentName);
        Assert.Equal("段階突破：国語（古文）・2020年度・第2回", result.StageTestSummary);
    }
}
