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
    public void MergeByReport_PutsAllProblemsBeforeAnswersAndKeepsFirstSeenReportOrder()
    {
        var ap1 = CreateMarked("ap1", [1, 2]);
        var aa1 = CreateMarked("aa1", [5]);
        var ap2 = CreateMarked("ap2", [3]);
        var aa2 = CreateMarked("aa2", [6]);
        var bp = CreateMarked("bp", [7]);
        var ba = CreateMarked("ba", [8]);
        // Interleaved successful jobs must still assemble one report packet at a time.
        var result = PdfBatchMergeService.MergeByReport([
            new("folder-A/report.pdf", new(ap1, aa1)),
            new("folder-B/report.pdf", new(bp, ba)),
            new("folder-A/report.pdf", new(ap2, aa2))], _directory, DateTimeOffset.Now);
        Assert.Equal(2, result.ReportCount);
        Assert.Equal(3, result.TestCount);
        Assert.Equal([1, 2, 3, 5, 6, 7, 8], ReadMarkers(result.PrintPdfPath));
    }

    [Fact]
    public void MergeByReport_KeepsWholeStageAndSkipsAbsentAnswerWithoutFabricatingOne()
    {
        var normal = CreateMarked("normal", [1]);
        var answer = CreateMarked("answer", [4]);
        var stage = CreateMarked("stage", [2, 3]);
        var problemOnly = CreateMarked("problem-only", [5]);
        var result = PdfBatchMergeService.MergeByReport([
            new("A", new(normal, answer)), new("A", new(stage, null)), new("B", new(problemOnly, null))],
            _directory, DateTimeOffset.Now);
        Assert.Equal([1, 2, 3, 4, 5], ReadMarkers(result.PrintPdfPath));
    }

    [Fact]
    public void MergeByReport_DuplexPreservesNewSheetStartsAtEveryPdfBoundary()
    {
        var ap = CreateMarked("ap", [1]);
        var aa = CreateMarked("aa", [2]);
        var bp = CreateMarked("bp", [3, 4]);
        var ba = CreateMarked("ba", [5]);
        var result = PdfBatchMergeService.MergeByReport([new("A", new(ap, aa)), new("B", new(bp, ba))],
            _directory, DateTimeOffset.Now, duplex: true);
        Assert.Equal([1, 0, 2, 0, 3, 4, 5], ReadMarkers(result.PrintPdfPath));
    }

    [Fact]
    public void MergeByReport_DoesNotOverwriteExistingPacketOrModifySources()
    {
        var source = CreateMarked("source", [1]);
        var before = File.ReadAllBytes(source);
        var jobs = new GeneratedReportTest[] { new("A", new(source, null)) };
        var now = DateTimeOffset.Now;
        var first = PdfBatchMergeService.MergeByReport(jobs, _directory, now);
        var original = File.ReadAllBytes(first.PrintPdfPath);
        var second = PdfBatchMergeService.MergeByReport(jobs, _directory, now);
        Assert.NotEqual(first.PrintPdfPath, second.PrintPdfPath);
        Assert.Equal(original, File.ReadAllBytes(first.PrintPdfPath));
        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Fact]
    public void MergeByReport_RejectsMissingAnswerAndEmptyInputBeforeCreatingBatchFolder()
    {
        var source = CreateMarked("source", [1]);
        Assert.Throws<ArgumentException>(() => PdfBatchMergeService.MergeByReport([], _directory, DateTimeOffset.Now));
        Assert.Throws<FileNotFoundException>(() => PdfBatchMergeService.MergeByReport([
            new("A", new(source, Path.Combine(_directory, "missing.pdf")))], _directory, DateTimeOffset.Now));
        Assert.False(Directory.Exists(Path.Combine(_directory, "一括印刷")));
    }

    private string CreateMarked(string name, int[] markers)
    {
        Directory.CreateDirectory(_directory);
        var path = Path.Combine(_directory, name + ".pdf");
        using var pdf = new PdfDocument();
        foreach (var marker in markers)
        {
            var page = pdf.AddPage();
            page.Width = PdfSharp.Drawing.XUnit.FromPoint(500 + marker);
            page.Rotate = 90;
        }
        pdf.Save(path);
        return path;
    }

    private static int[] ReadMarkers(string path)
    {
        using var pdf = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        return pdf.Pages.Cast<PdfPage>().Select(page => page.Rotate == 90 ? (int)page.Width.Point - 500 : 0).ToArray();
    }

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
        Assert.Equal(3, ReadPageCount(result.AnswerPdfPath!));
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
