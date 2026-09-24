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

    private readonly OcrEngine _ocrEngine;

    public WindowsReportOcrService()
    {
        _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ja-JP"))
            ?? OcrEngine.TryCreateFromUserProfileLanguages()
            ?? throw new InvalidOperationException(
                "Windowsの日本語OCRを使用できません。Windowsの言語設定で日本語を追加してください。");
    }

    public async Task<RecognizedReport> RecognizeAsync(
        string pdfPath,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var storageFile = await StorageFile.GetFileFromPathAsync(pdfPath);
        var document = await PdfDocument.LoadFromFileAsync(storageFile);
        var pages = new List<RecognizedReportPage>((int)document.PageCount);

        for (uint index = 0; index < document.PageCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report($"{Path.GetFileName(pdfPath)}：{index + 1}/{document.PageCount}ページ");

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
        var result = await _ocrEngine.RecognizeAsync(bitmap);
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
