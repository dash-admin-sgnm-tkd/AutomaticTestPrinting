using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using System.IO;
using Windows.Data.Pdf;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

namespace AutomaticTestPrinting.App.Services;

public sealed class WindowsReportOcrService
{
    private static readonly string[] OrientationKeywords =
    [
        "テスト", "作成", "依頼", "生徒", "次回", "宿題", "段階", "科目", "範囲", "特訓"
    ];

    private readonly Lazy<OcrEngine> _ocrEngine = new(() =>
        OcrEngine.TryCreateFromLanguage(new Language("ja-JP"))
            ?? OcrEngine.TryCreateFromUserProfileLanguages()
            ?? throw new InvalidOperationException(
                "Windowsの日本語OCRを使用できません。Windowsの言語設定で日本語を追加してください。"));

    public async Task<RecognizedReport> RecognizeAsync(
        string pdfPath,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var storageFile = await StorageFile.GetFileFromPathAsync(pdfPath);
        var document = await PdfDocument.LoadFromFileAsync(storageFile);
        var pages = new List<RecognizedReportPage>((int)document.PageCount);
        IReadOnlyList<RecognizedReportPage?> directPages;
        try
        {
            directPages = await Task.Run(() => PdfReportTextExtractionService.Extract(pdfPath, cancellationToken), cancellationToken);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            // A PDF unsupported by the text reader may still be rendered by Windows.
            progress?.Report($"文字情報を使用できないためOCRで読み取ります（{exception.GetType().Name}）");
            directPages = [];
        }

        for (uint index = 0; index < document.PageCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report($"{Path.GetFileName(pdfPath)}：{index + 1}/{document.PageCount}ページ");

            if (index < directPages.Count && directPages[(int)index] is { } direct)
            {
                pages.Add(direct);
                progress?.Report($"{Path.GetFileName(pdfPath)}：{index + 1}ページをPDF文字情報から読み取りました");
                continue;
            }

            using var page = document.GetPage(index);
            using var stream = new InMemoryRandomAccessStream();
            var width = 2200u;
            var height = (uint)Math.Round(width * page.Size.Height / page.Size.Width);
            var options = new PdfPageRenderOptions
            {
                DestinationWidth = width,
                DestinationHeight = height
            };

            await page.RenderToStreamAsync(stream, options);
            var upright = await RecognizeRotationAsync(stream, BitmapRotation.None);
            var upsideDown = await RecognizeRotationAsync(stream, BitmapRotation.Clockwise180Degrees);
            var selected = Score(upsideDown.Text) > Score(upright.Text)
                ? CreatePageResult((int)index + 1, 180, upsideDown)
                : CreatePageResult((int)index + 1, 0, upright);
            pages.Add(selected);
        }

        return ReportTextParser.Parse(pdfPath, pages);
    }

    private async Task<OcrCandidate> RecognizeRotationAsync(
        InMemoryRandomAccessStream stream,
        BitmapRotation rotation)
    {
        stream.Seek(0);
        var decoder = await BitmapDecoder.CreateAsync(stream);
        var transform = new BitmapTransform { Rotation = rotation };
        using var bitmap = await decoder.GetSoftwareBitmapAsync(
            BitmapPixelFormat.Bgra8,
            BitmapAlphaMode.Premultiplied,
            transform,
            ExifOrientationMode.IgnoreExifOrientation,
            ColorManagementMode.DoNotColorManage);
        var result = await _ocrEngine.Value.RecognizeAsync(bitmap);
        var words = result.Lines
            .SelectMany((line, lineIndex) => line.Words.Select(word => new RecognizedWord(
                word.Text,
                word.BoundingRect.X,
                word.BoundingRect.Y,
                word.BoundingRect.Width,
                word.BoundingRect.Height,
                lineIndex)))
            .ToArray();
        return new OcrCandidate(result.Text, bitmap.PixelWidth, bitmap.PixelHeight, words);
    }

    private static RecognizedReportPage CreatePageResult(
        int pageNumber,
        int rotationDegrees,
        OcrCandidate candidate) => new(
            pageNumber,
            rotationDegrees,
            candidate.Text,
            candidate.PixelWidth,
            candidate.PixelHeight,
            candidate.Words);

    private static int Score(string text)
    {
        var score = text.Length / 40;
        foreach (var keyword in OrientationKeywords)
        {
            score += CountOccurrences(text, keyword) * 25;
        }

        return score;
    }

    private static int CountOccurrences(string text, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private sealed record OcrCandidate(
        string Text,
        int PixelWidth,
        int PixelHeight,
        IReadOnlyList<RecognizedWord> Words);
}
