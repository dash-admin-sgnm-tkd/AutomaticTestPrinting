using AutomaticTestPrinting.Core.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Core;

namespace AutomaticTestPrinting.Core.Services;

/// <summary>Read the text layer with coordinates. Unusable pages remain null for OCR fallback.</summary>
public static class PdfReportTextExtractionService
{
    public static IReadOnlyList<RecognizedReportPage?> Extract(string path, CancellationToken cancellationToken = default)
    {
        using var document = PdfDocument.Open(path);
        var pages = new List<RecognizedReportPage?>(document.NumberOfPages);
        foreach (var page in document.GetPages())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var scale = 2200 / page.Width;
            var lines = new List<List<UglyToad.PdfPig.Content.Letter>>();
            foreach (var letter in page.Letters.OrderByDescending(letter => letter.StartBaseLine.Y))
            {
                var line = lines.LastOrDefault();
                if (line is null || Math.Abs(line[0].StartBaseLine.Y - letter.StartBaseLine.Y) > 1.5)
                    lines.Add([letter]);
                else line.Add(letter);
            }
            var words = lines.SelectMany((line, index) => line.OrderBy(letter => letter.BoundingBox.Left)
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Select(letter => new RecognizedWord(letter.Value,
                    letter.BoundingBox.Left * scale, (page.Height - letter.BoundingBox.Top) * scale,
                    letter.BoundingBox.Width * scale, letter.BoundingBox.Height * scale, index))).ToArray();
            var text = string.Join('\n', lines.Select(line => string.Concat(line.OrderBy(letter => letter.BoundingBox.Left).Select(letter => letter.Value))));
            var tableLines = page.Paths.Where(path => path.IsStroked || path.IsFilled)
                .SelectMany(path => path).SelectMany(subpath => subpath.Commands.OfType<PdfSubpath.Line>())
                .Where(line => (Math.Abs(line.From.Y - line.To.Y) < 0.5 && Math.Abs(line.From.X - line.To.X) > 10) ||
                               (Math.Abs(line.From.X - line.To.X) < 0.5 && Math.Abs(line.From.Y - line.To.Y) > 10))
                .Select(line => new RecognizedTableLine(line.From.X * scale, (page.Height - line.From.Y) * scale,
                    line.To.X * scale, (page.Height - line.To.Y) * scale)).ToArray();
            var candidate = new RecognizedReportPage(page.Number, 0, text, 2200,
                (int)Math.Round(page.Height * scale), words, "PDF文字抽出", tableLines);
            pages.Add(IsUsable(candidate) ? candidate : null);
        }
        return pages;
    }

    public static bool IsUsable(RecognizedReportPage page)
    {
        var compact = string.Concat(page.Text.Where(character => !char.IsWhiteSpace(character)));
        if (compact.Length < 40 || compact.Contains('\uFFFD') || compact.Contains('\0') || page.RecognizedWords.Count < 20)
            return false;
        if (!compact.Contains("生徒名", StringComparison.Ordinal) &&
            !compact.Contains("次回までの宿題", StringComparison.Ordinal) &&
            !compact.Contains("次回の内容", StringComparison.Ordinal)) return false;
        // A label-only or incomplete text layer must not hide the table visible in an image.
        if (compact.Contains("次回までの宿題", StringComparison.Ordinal))
            return compact.Contains("テスト作成依頼", StringComparison.Ordinal) &&
                ReportTableExtractor.ExtractNormalTestRequests([page]).Count > 0;
        return true;
    }
}
