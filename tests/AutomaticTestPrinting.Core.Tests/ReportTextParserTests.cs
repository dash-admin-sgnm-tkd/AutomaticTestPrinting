using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReportTextParserTests
{
    [Theory]
    [InlineData("100題", 100)]
    [InlineData("25問", 25)]
    [InlineData("25", 25)]
    public void Parse_AcceptsMisrecognizedCountSuffixWithinRequestRow(string count, int expected)
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("高校英語", 380, 50, 80, 20, 1),
            new RecognizedWord("教材A", 380, 85, 80, 20, 2),
            new RecognizedWord("月日", 10, 150, 40, 20, 3),
            new RecognizedWord("1-200", 380, 150, 80, 20, 4),
            new RecognizedWord("テスト作成依頼", 10, 300, 140, 20, 5),
            new RecognizedWord(count, 400, 302, 40, 20, 6)
        };
        var report = ReportTextParser.Parse("report.pdf", [new(1, 0, "次回までの宿題 テスト作成依頼", 1000, 500, words)]);
        Assert.Equal(expected, Assert.Single(report.NormalTestRequests!).QuestionCount);
        Assert.Empty(report.NormalTestReadWarning);
    }

    [Fact]
    public void Parse_MissingCountStaysAsUnconfirmedCandidateAndDoesNotCopyOtherCount()
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("高校英語", 380, 50, 80, 20, 1),
            new RecognizedWord("高校英語", 580, 50, 80, 20, 1),
            new RecognizedWord("教材A", 380, 85, 80, 20, 2),
            new RecognizedWord("教材B", 580, 85, 80, 20, 3),
            new RecognizedWord("月日", 10, 150, 40, 20, 4),
            new RecognizedWord("1-200", 380, 150, 80, 20, 5),
            new RecognizedWord("1-20", 580, 150, 80, 20, 6),
            new RecognizedWord("テスト作成依頼", 10, 300, 140, 20, 7),
            new RecognizedWord("25題", 400, 302, 40, 20, 7)
        };
        var report = ReportTextParser.Parse("report.pdf", [new(1, 0, "次回までの宿題 テスト作成依頼", 1000, 500, words)]);
        Assert.Equal(2, report.NormalTestRequests!.Count);
        Assert.Equal(25, report.NormalTestRequests[0].QuestionCount);
        Assert.Equal("教材B", report.NormalTestRequests[1].MaterialName);
        Assert.Equal(0, report.NormalTestRequests[1].QuestionCount);
        Assert.True(report.NormalTestRequests[1].RequiresReview);
    }

    [Fact]
    public void Parse_ZeroRequestsWithVisibleRequestLabelHasExplicitWarning()
    {
        var words = new[] { new RecognizedWord("次回までの宿題", 10, 10, 100, 20, 0),
            new RecognizedWord("テスト作成依頼", 10, 300, 140, 20, 1) };
        var report = ReportTextParser.Parse("report.pdf", [new(1, 0, "次回までの宿題 テスト作成依頼", 1000, 500, words)]);
        Assert.Empty(report.NormalTestRequests!);
        Assert.Contains("要確認", report.NormalTestSummary);
        Assert.Contains("依頼なしとは判断せず", report.NormalTestReadWarning);
    }

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
        Assert.NotNull(result.StageTestRequest);
        Assert.Equal("段階突破 国語（古文） 2020年度 日大 第2回", result.StageTestRequest.MaterialName);
        Assert.Equal("1-1", result.StageTestRequest.Range);
        Assert.Single(result.TestRequests);
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

    [Fact]
    public void Parse_DoesNotGuessTarget1900RangeDigits()
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("周校英語", 370, 50, 100, 20, 1),
            new RecognizedWord("[六]英単語ターゲット1900", 350, 85, 160, 20, 2),
            new RecognizedWord("月日", 10, 150, 40, 20, 3),
            new RecognizedWord("曜日", 60, 150, 40, 20, 4),
            new RecognizedWord("1901-1900", 370, 150, 100, 20, 5),
            new RecognizedWord("テスト作成依頼", 10, 300, 160, 20, 6),
            new RecognizedWord("50", 405, 302, 30, 20, 7)
        };
        var pages = new[]
        {
            new RecognizedReportPage(
                1,
                0,
                "生徒名:テスト生徒 次回までの宿題 テスト作成依頼",
                1000,
                500,
                words)
        };

        var result = ReportTextParser.Parse("report.pdf", pages);

        var request = Assert.Single(result.TestRequests);
        Assert.Equal("[六]英単語ターゲット1900", request.MaterialName);
        Assert.Equal("1901-1900", request.Range);
        Assert.Equal(50, request.QuestionCount);
    }

    [Theory]
    [InlineData("ｲ1101-1900", "1101-1900")]
    [InlineData("範囲ｲ1101～1900問", "1101-1900")]
    [InlineData("１１０１－１９００", "1101-1900")]
    [InlineData("1-900", "1-900")]
    public void Parse_RemovesOcrNoiseAroundRangeNumbers(string ocrRange, string expectedRange)
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("英単語ターゲット1900", 350, 85, 160, 20, 1),
            new RecognizedWord("月日", 10, 150, 40, 20, 2),
            new RecognizedWord(ocrRange, 370, 150, 120, 20, 3),
            new RecognizedWord("テスト作成依頼", 10, 300, 160, 20, 4),
            new RecognizedWord("50", 405, 302, 30, 20, 5)
        };
        var pages = new[]
        {
            new RecognizedReportPage(
                1,
                0,
                "生徒名:テスト生徒 次回までの宿題 テスト作成依頼",
                1000,
                500,
                words)
        };

        var result = ReportTextParser.Parse("report.pdf", pages);

        var request = Assert.Single(result.TestRequests);
        Assert.Equal(expectedRange, request.Range);
        Assert.Equal(50, request.QuestionCount);
    }
}
