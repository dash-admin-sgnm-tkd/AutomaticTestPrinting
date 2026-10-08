using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfReportTextExtractionTests
{
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
