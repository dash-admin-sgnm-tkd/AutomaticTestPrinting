using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfReportTextExtractionTests
{
    [Fact]
    public void ExtractTable_UsesRulesWhenMonthLabelIsBottomAlignedWithWrappedRange()
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("高校英語", 280, 50, 80, 20, 1),
            new RecognizedWord("大学入試はじめの", 260, 80, 130, 20, 2),
            new RecognizedWord("英文読解ドリル", 265, 105, 120, 20, 3),
            new RecognizedWord("関正生のTheRules", 440, 80, 140, 20, 2),
            new RecognizedWord("英語長文問題集1入試基礎", 430, 130, 150, 20, 4),
            new RecognizedWord("Vintage4thEdition", 625, 80, 150, 20, 2),
            new RecognizedWord("c1-1~入試実", 260, 180, 110, 20, 5),
            new RecognizedWord("践演習2", 280, 205, 70, 20, 6),
            new RecognizedWord("1-6", 470, 180, 50, 20, 5),
            new RecognizedWord("1,2章", 660, 180, 60, 20, 5),
            new RecognizedWord("月日", 10, 205, 40, 20, 6),
            new RecognizedWord("テスト作成依頼", 10, 400, 140, 20, 7),
            new RecognizedWord("10", 310, 400, 30, 20, 7),
            new RecognizedWord("1", 495, 400, 20, 20, 7),
            new RecognizedWord("10", 680, 400, 30, 20, 7)
        };
        var rules = new[]
        {
            new RecognizedTableLine(220, 170, 800, 170),
            new RecognizedTableLine(220, 235, 800, 235),
            new RecognizedTableLine(220, 50, 220, 430),
            new RecognizedTableLine(410, 50, 410, 430),
            new RecognizedTableLine(600, 50, 600, 430),
            new RecognizedTableLine(800, 50, 800, 430)
        };
        var page = new RecognizedReportPage(2, 0, "", 1000, 500, words, "PDF文字抽出", rules);
        var requests = ReportTableExtractor.ExtractNormalTestRequests([page]);
        Assert.Equal(3, requests.Count);
        Assert.Equal("大学入試はじめの英文読解ドリル", requests[0].MaterialName);
        Assert.Equal("c1-1~入試実践演習2", requests[0].Range);
        Assert.Equal(10, requests[0].QuestionCount);
        Assert.Equal("関正生のTheRules英語長文問題集1入試基礎", requests[1].MaterialName);
        Assert.Equal("1-6", requests[1].Range);
        Assert.Equal(1, requests[1].QuestionCount);
        Assert.Equal("Vintage4thEdition", requests[2].MaterialName);
        Assert.Equal("1,2章", requests[2].Range);
        Assert.Equal(10, requests[2].QuestionCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("生徒名")]
    [InlineData("生徒名\uFFFD次回の内容と長い文字列ですが文字情報が破損しています。文字情報が破損しています。文字情報が破損しています。")]
    public void IsUsable_EmptySparseOrBrokenLayerFallsBackToOcr(string text)
    {
        Assert.False(PdfReportTextExtractionService.IsUsable(new(1, 0, text, 2200, 1556, [], "PDF文字抽出")));
    }

    [Fact]
    public void ExtractTable_DirectTextKeepsNumbersAndSkipsTrulyBlankRequestColumn()
    {
        var words = new[]
        {
            new RecognizedWord("次回までの宿題", 10, 10, 140, 20, 0),
            new RecognizedWord("高校英語", 380, 50, 80, 20, 1),
            new RecognizedWord("高校英語", 580, 50, 80, 20, 1),
            new RecognizedWord("竹岡の英語長文SUPREMACY至高の20題", 360, 85, 120, 20, 2),
            new RecognizedWord("英検過去問", 580, 85, 80, 20, 3),
            new RecognizedWord("月日", 10, 150, 40, 20, 4),
            new RecognizedWord("11-15", 380, 150, 80, 20, 5),
            new RecognizedWord("0-0", 580, 150, 80, 20, 6),
            new RecognizedWord("テスト作成依頼", 10, 300, 140, 20, 7),
            new RecognizedWord("1題", 400, 302, 40, 20, 7)
        };
        var request = Assert.Single(ReportTableExtractor.ExtractNormalTestRequests([new(1, 0, "", 1000, 500, words, "PDF文字抽出")]));
        Assert.Equal("竹岡の英語長文SUPREMACY至高の20題", request.MaterialName);
        Assert.Equal("11-15", request.Range);
        Assert.Equal(1, request.QuestionCount);
    }
}
