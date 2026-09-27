using AutomaticTestPrinting.Core.Models;
using System.Globalization;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Services;

public static class PdfBatchMergeService
{
    public static BatchPdfGenerationResult Merge(
        IReadOnlyList<ExcelPdfGenerationResult> generatedTests,
        string outputFolder,
        DateTimeOffset createdAt)
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
            EnsureSourceExists(test.AnswerPdfPath);
        }

        var batchFolder = Path.Combine(outputFolder, "一括印刷");
        Directory.CreateDirectory(batchFolder);
        var suffix = createdAt.ToString("yyyyMMdd_HHmmssfff", CultureInfo.InvariantCulture);
        var problemPath = CreateAvailablePath(batchFolder, $"確認テスト_問題一括_{suffix}.pdf");
        var answerPath = CreateAvailablePath(batchFolder, $"確認テスト_解答一括_{suffix}.pdf");

        var problemTemporaryPath = problemPath + ".tmp";
        var answerTemporaryPath = answerPath + ".tmp";
        try
        {
            MergeFiles(generatedTests.Select(test => test.ProblemPdfPath), problemTemporaryPath);
            MergeFiles(generatedTests.Select(test => test.AnswerPdfPath), answerTemporaryPath);
            File.Move(problemTemporaryPath, problemPath);
            File.Move(answerTemporaryPath, answerPath);
            return new BatchPdfGenerationResult(problemPath, answerPath, generatedTests.Count);
        }
        catch
        {
            TryDelete(problemTemporaryPath);
            TryDelete(answerTemporaryPath);
            TryDelete(problemPath);
            TryDelete(answerPath);
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
