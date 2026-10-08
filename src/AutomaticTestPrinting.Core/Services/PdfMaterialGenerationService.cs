using System.Globalization;
using AutomaticTestPrinting.Core.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Services;

public static class PdfMaterialGenerationService
{
    // Numbers refer to PDF viewer pages, not printed page/question numbers.
    public static IReadOnlyList<int> ParsePages(string pages, int pageCount)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(pageCount, 1);
        ArgumentNullException.ThrowIfNull(pages);
        if (string.IsNullOrWhiteSpace(pages))
        {
            return Enumerable.Range(1, pageCount).ToArray();
        }

        var selected = new List<int>();
        var seen = new HashSet<int>();
        foreach (var token in pages.Replace('，', ',').Split(','))
        {
            var bounds = token.Trim().Split('-');
            if (bounds.Length is < 1 or > 2 ||
                !int.TryParse(bounds[0].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var start) ||
                (bounds.Length == 2 && !int.TryParse(bounds[1].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out _)))
            {
                throw new ArgumentException("ページは「1-3,5」の形式で入力してください。空欄は全ページです。", nameof(pages));
            }
            var end = bounds.Length == 2 ? int.Parse(bounds[1].Trim(), CultureInfo.InvariantCulture) : start;
            if (start < 1 || end < start || end > pageCount)
            {
                throw new ArgumentException($"ページは1-{pageCount}内で、開始≦終了になるよう指定してください：{token}", nameof(pages));
            }
            for (var page = start; page <= end; page++)
            {
                if (!seen.Add(page))
                {
                    throw new ArgumentException($"ページ {page} が重複しています。", nameof(pages));
                }
                selected.Add(page);
            }
        }
        return selected;
    }

    public static int GetPageCount(string path)
    {
        using var input = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        return input.PageCount;
    }

    public static PdfMaterialBatchResult Generate(
        IReadOnlyList<PdfMaterialRequest> requests,
        string outputFolder,
        bool duplex)
    {
        ArgumentNullException.ThrowIfNull(requests);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFolder);
        if (requests.Count == 0)
        {
            throw new ArgumentException("作成するPDFを追加してください。", nameof(requests));
        }
        // A new run folder prevents replacement of either original or previous PDFs.
        var runFolder = Path.Combine(outputFolder, $"PDF教材_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}");
        Directory.CreateDirectory(runFolder);
        var outputs = new List<PdfMaterialOutput>();
        var errors = new List<ErrorLogEntry>();
        for (var index = 0; index < requests.Count; index++)
        {
            var request = requests[index];
            var destination = Path.Combine(runFolder,
                $"{index + 1:D3}_{SafeName(request.StudentName)}_{SafeName(request.MaterialName)}_{(request.IsAnswer ? "解答" : "問題")}.pdf");
            try
            {
                using var input = PdfReader.Open(request.SourcePath, PdfDocumentOpenMode.Import);
                var pages = ParsePages(request.Pages, input.PageCount);
                using var output = new PdfDocument();
                foreach (var page in pages)
                {
                    output.AddPage(input.Pages[page - 1]);
                }
                output.Save(destination);
                outputs.Add(new PdfMaterialOutput(request, destination, pages.Count));
            }
            catch (Exception exception) when (exception is not OutOfMemoryException)
            {
                errors.Add(new ErrorLogEntry(DateTimeOffset.Now, "PDF教材・ページ抽出",
                    request.StudentName, Path.GetFileName(request.SourcePath), index + 1,
                    request.MaterialName, string.IsNullOrWhiteSpace(request.Pages) ? "全ページ" : request.Pages,
                    string.Empty, exception.GetType().FullName ?? exception.GetType().Name,
                    $"0x{exception.HResult:X8}", exception.Message,
                    $"元PDF：{request.SourcePath}\n保存先：{destination}\n{exception}"));
            }
        }

        string? MergeRole(bool answer)
        {
            var files = outputs.Where(output => output.Request.IsAnswer == answer).Select(output => output.Path).ToArray();
            if (files.Length == 0)
            {
                return null;
            }
            var destination = Path.Combine(runFolder, answer ? "解答一括.pdf" : "問題一括.pdf");
            MergeFiles(files, destination, duplex);
            return destination;
        }

        return new PdfMaterialBatchResult(runFolder, outputs, MergeRole(false), MergeRole(true), errors);
    }

    public static void MergeFiles(IReadOnlyList<string> paths, string destination, bool duplex)
    {
        ArgumentNullException.ThrowIfNull(paths);
        ArgumentException.ThrowIfNullOrWhiteSpace(destination);
        if (paths.Count == 0)
        {
            throw new ArgumentException("結合するPDFがありません。", nameof(paths));
        }
        if (File.Exists(destination))
        {
            throw new IOException($"既存ファイルは上書きできません：{destination}");
        }
        using var output = new PdfDocument();
        for (var fileIndex = 0; fileIndex < paths.Count; fileIndex++)
        {
            using var input = PdfReader.Open(paths[fileIndex], PdfDocumentOpenMode.Import);
            if (input.PageCount == 0)
            {
                throw new InvalidOperationException($"ページがありません：{paths[fileIndex]}");
            }
            foreach (var page in input.Pages)
            {
                output.AddPage(page);
            }
            // Start the next file on a new sheet, even if this one has an odd page count.
            if (duplex && fileIndex < paths.Count - 1 && output.PageCount % 2 != 0)
            {
                var last = output.Pages[output.PageCount - 1];
                var blank = output.AddPage();
                blank.Width = last.Width;
                blank.Height = last.Height;
            }
        }
        output.Save(destination);
    }

    private static string SafeName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var name = string.Concat(value.Select(character => invalid.Contains(character) || char.IsControl(character) ? '_' : character));
        name = name.Trim().TrimEnd('.');
        return name.Length == 0 ? "未指定" : name[..Math.Min(name.Length, 40)];
    }
}
