using AutomaticTestPrinting.Core.Models;
using System.Globalization;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Services;

public static class PdfBatchMergeService
{
    public static ReportBatchPdfGenerationResult MergeByReport(
        IReadOnlyList<GeneratedReportTest> generatedTests, string outputFolder, DateTimeOffset createdAt, bool duplex = false)
    {
        ArgumentNullException.ThrowIfNull(generatedTests);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFolder);
        if (generatedTests.Count == 0) throw new ArgumentException("結合するテストがありません。", nameof(generatedTests));
        foreach (var test in generatedTests)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(test.SourceReportPath);
            EnsureSourceExists(test.Pdf.ProblemPdfPath);
            if (test.Pdf.AnswerPdfPath is not null) EnsureSourceExists(test.Pdf.AnswerPdfPath);
        }
        // Keep report order and test order; never group solely by student or report filename.
        var reports = generatedTests.GroupBy(test => test.SourceReportPath, StringComparer.OrdinalIgnoreCase).ToArray();
        var paths = reports.SelectMany(report => report.Select(test => test.Pdf.ProblemPdfPath)
            .Concat(report.Where(test => test.Pdf.AnswerPdfPath is not null).Select(test => test.Pdf.AnswerPdfPath!))).ToArray();
        var folder = Path.Combine(outputFolder, "一括印刷");
        Directory.CreateDirectory(folder);
        var suffix = createdAt.ToString("yyyyMMdd_HHmmssfff", CultureInfo.InvariantCulture);
        var destination = CreateAvailablePath(folder, $"確認テスト_レポート順_問題解答_{suffix}.pdf");
        var temporary = destination + ".tmp";
        try
        {
            PdfMaterialGenerationService.MergeFiles(paths, temporary, duplex);
            File.Move(temporary, destination);
            return new(destination, generatedTests.Count, reports.Length);
        }
        catch
        {
            TryDelete(temporary);
            throw;
        }
    }

    public static BatchPdfGenerationResult Merge(
        IReadOnlyList<ExcelPdfGenerationResult> generatedTests,
        string outputFolder,
        DateTimeOffset createdAt,
        bool duplex = false)
    {
        ArgumentNullException.ThrowIfNull(generatedTests);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFolder);

        if (generatedTests.Count == 0)
        {
            throw new ArgumentException("一括PDFにまとめるテストがありません。", nameof(generatedTests));
        }

        foreach (var test in generatedTests)
        {
            EnsureSourceExists(test.ProblemPdfPath);
            if (test.AnswerPdfPath is not null) EnsureSourceExists(test.AnswerPdfPath);
        }

        var batchFolder = Path.Combine(outputFolder, "一括印刷");
        Directory.CreateDirectory(batchFolder);
        var suffix = createdAt.ToString("yyyyMMdd_HHmmssfff", CultureInfo.InvariantCulture);
        var problemPath = CreateAvailablePath(batchFolder, $"確認テスト_問題一括_{suffix}.pdf");
        var answers = generatedTests.Where(test => test.AnswerPdfPath is not null).Select(test => test.AnswerPdfPath!).ToArray();
        var answerPath = answers.Length > 0 ? CreateAvailablePath(batchFolder, $"確認テスト_解答一括_{suffix}.pdf") : null;

        var problemTemporaryPath = problemPath + ".tmp";
        var answerTemporaryPath = answerPath is null ? null : answerPath + ".tmp";
        try
        {
            PdfMaterialGenerationService.MergeFiles(generatedTests.Select(test => test.ProblemPdfPath).ToArray(), problemTemporaryPath, duplex);
            if (answerTemporaryPath is not null) PdfMaterialGenerationService.MergeFiles(answers, answerTemporaryPath, duplex);
            File.Move(problemTemporaryPath, problemPath);
            if (answerTemporaryPath is not null) File.Move(answerTemporaryPath, answerPath!);
            return new BatchPdfGenerationResult(problemPath, answerPath, generatedTests.Count);
        }
        catch
        {
            TryDelete(problemTemporaryPath);
            if (answerTemporaryPath is not null) TryDelete(answerTemporaryPath);
            TryDelete(problemPath);
            if (answerPath is not null) TryDelete(answerPath);
            throw;
        }
    }

    private static void MergeFiles(IEnumerable<string> sourcePaths, string destinationPath)
    {
        using var output = new PdfDocument();
        foreach (var sourcePath in sourcePaths)
        {
            using var input = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Import);
            for (var pageIndex = 0; pageIndex < input.PageCount; pageIndex++)
            {
                output.AddPage(input.Pages[pageIndex]);
            }
        }

        if (output.PageCount == 0)
        {
            throw new InvalidOperationException("一括PDFに追加できるページがありませんでした。");
        }

        output.Save(destinationPath);
    }

    private static void EnsureSourceExists(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length == 0)
        {
            throw new FileNotFoundException("結合するPDFが見つかりません。", path);
        }
    }

    private static string CreateAvailablePath(string folder, string fileName)
    {
        var path = Path.Combine(folder, fileName);
        if (!File.Exists(path))
        {
            return path;
        }

        var stem = Path.GetFileNameWithoutExtension(fileName);
        for (var sequence = 2; sequence < int.MaxValue; sequence++)
        {
            path = Path.Combine(folder, $"{stem}_{sequence}.pdf");
            if (!File.Exists(path))
            {
                return path;
            }
        }

        throw new IOException("一括PDFの保存先ファイル名を決められませんでした。");
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
