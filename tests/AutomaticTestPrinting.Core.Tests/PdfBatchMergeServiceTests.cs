using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfBatchMergeServiceTests : IDisposable
{
    private readonly string _directory = Path.Combine(
        Path.GetTempPath(),
        $"AutomaticTestPrinting-PdfMerge-{Guid.NewGuid():N}");

    [Fact]
    public void Merge_CreatesProblemAndAnswerPdfInInputOrder()
    {
        Directory.CreateDirectory(_directory);
        var first = CreateTestResult("first", problemPages: 1, answerPages: 2);
        var second = CreateTestResult("second", problemPages: 3, answerPages: 1);

        var result = PdfBatchMergeService.Merge(
            [first, second],
            _directory,
            new DateTimeOffset(2026, 9, 27, 9, 30, 0, TimeSpan.FromHours(9)));

        Assert.Equal(2, result.TestCount);
        Assert.Equal(4, ReadPageCount(result.ProblemPdfPath));
        Assert.Equal(3, ReadPageCount(result.AnswerPdfPath));
        Assert.Contains("一括印刷", result.ProblemPdfPath);
        Assert.Contains("確認テスト_問題一括_20260927_093000000.pdf", result.ProblemPdfPath);
    }

    [Fact]
    public void Merge_RejectsMissingSourcePdfWithoutLeavingBatchFiles()
    {
        Directory.CreateDirectory(_directory);
        var result = new ExcelPdfGenerationResult(
            Path.Combine(_directory, "missing-problem.pdf"),
            Path.Combine(_directory, "missing-answer.pdf"));

        Assert.Throws<FileNotFoundException>(() => PdfBatchMergeService.Merge(
            [result],
            _directory,
            DateTimeOffset.Now));
        Assert.False(Directory.Exists(Path.Combine(_directory, "一括印刷")));
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory))
        {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private ExcelPdfGenerationResult CreateTestResult(
        string name,
        int problemPages,
        int answerPages)
    {
        var problemPath = Path.Combine(_directory, $"{name}-problem.pdf");
        var answerPath = Path.Combine(_directory, $"{name}-answer.pdf");
        CreatePdf(problemPath, problemPages);
        CreatePdf(answerPath, answerPages);
        return new ExcelPdfGenerationResult(problemPath, answerPath);
    }

    private static void CreatePdf(string path, int pageCount)
    {
        using var document = new PdfDocument();
        for (var index = 0; index < pageCount; index++)
        {
            document.AddPage();
        }
        document.Save(path);
    }

    private static int ReadPageCount(string path)
    {
        using var document = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        return document.PageCount;
    }
}
