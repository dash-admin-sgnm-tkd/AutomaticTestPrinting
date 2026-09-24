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
        Assert.Equal("依頼欄あり（問題数は画面で確認してください）", result.NormalTestSummary);
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

    [Fact]
    public void Parse_LinksCountToMaterialAndRangeByColumnPosition()
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("高校英語", 370, 50, 100, 20, 1),
            new RecognizedWord("英単語ターゲット1900", 350, 85, 140, 20, 2),
            new RecognizedWord("月日", 10, 150, 40, 20, 3),
            new RecognizedWord("曜日", 60, 150, 40, 20, 4),
            new RecognizedWord("1-1900", 380, 150, 80, 20, 5),
            new RecognizedWord("9/20", 10, 200, 50, 20, 6),
            new RecognizedWord("テスト作成依頼", 10, 300, 160, 20, 7),
            new RecognizedWord("50", 405, 302, 30, 20, 8)
        };
        var pages = new[]
        {
            new RecognizedReportPage(
                1,
                0,
                "生徒名:山田太郎 講師名:講師 次回までの宿題 テスト作成依頼",
                1000,
                500,
                words)
        };

        var result = ReportTextParser.Parse("report.pdf", pages);

        var request = Assert.Single(result.TestRequests);
        Assert.Equal("英単語ターゲット1900", request.MaterialName);
        Assert.Equal("1-1900", request.Range);
        Assert.Equal(50, request.QuestionCount);
        Assert.Equal("通常テスト候補：1件（原本確認）", result.NormalTestSummary);
    }
}
