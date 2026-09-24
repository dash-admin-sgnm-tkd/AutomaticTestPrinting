using System.Globalization;
using System.Text.RegularExpressions;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static partial class ReportTableExtractor
{
    public static IReadOnlyList<NormalTestRequestCandidate> ExtractNormalTestRequests(
        IReadOnlyList<RecognizedReportPage> pages)
    {
        foreach (var page in pages.Reverse())
        {
            var lines = BuildLines(page);
            var marker = lines.LastOrDefault(line => line.Text.Contains("次回までの宿題", StringComparison.Ordinal));
            if (marker is null)
            {
                continue;
            }

            var requestLine = lines.LastOrDefault(line =>
                line.Top > marker.Top && line.Text.Contains("テスト作成依頼", StringComparison.Ordinal));
            if (requestLine is null)
            {
                continue;
            }

            var rangeAnchor = lines.FirstOrDefault(line =>
                line.Top > marker.Top &&
                line.Top < requestLine.Top &&
                line.Left < page.PixelWidth * 0.15 &&
                (line.Text.Contains("月日", StringComparison.Ordinal) ||
                 line.Text.Contains("曜日", StringComparison.Ordinal)));
            if (rangeAnchor is null)
            {
                continue;
            }

            var verticalTolerance = page.PixelHeight * 0.04;
            var countLines = lines
                .Where(line =>
                    line.Left > requestLine.Right + 20 &&
                    Math.Abs(line.CenterY - requestLine.CenterY) <= verticalTolerance)
                .Select(line => new { Line = line, Match = CountRegex().Match(line.Text) })
                .Where(item => item.Match.Success)
                .Select(item => new
                {
                    item.Line,
                    Count = int.Parse(item.Match.Groups[1].Value, CultureInfo.InvariantCulture)
                })
                .Where(item => item.Count is > 0 and <= 200)
                .OrderBy(item => item.Line.CenterX)
                .ToArray();

            var results = new List<NormalTestRequestCandidate>(countLines.Length);
            foreach (var item in countLines)
            {
                var halfColumnWidth = page.PixelWidth * 0.055;
                var material = ExtractMaterial(
                    page,
                    item.Line.CenterX,
                    halfColumnWidth,
                    marker.Bottom,
                    rangeAnchor.Top);
                var range = ExtractRange(
                    page,
                    item.Line.CenterX,
                    halfColumnWidth,
                    rangeAnchor.CenterY);

                if (material.Contains("段階突破", StringComparison.Ordinal))
                {
                    continue;
                }

                var normalizedRange = string.IsNullOrWhiteSpace(range)
                    ? "範囲を確認"
                    : RepairRangeForMaterial(NormalizeRange(range), material);

                results.Add(new NormalTestRequestCandidate(
                    string.IsNullOrWhiteSpace(material) ? "教材名を確認" : material,
                    normalizedRange,
                    item.Count,
                    page.PageNumber,
                    true));
            }

            return results;
        }

        return [];
    }

    private static string ExtractMaterial(
        RecognizedReportPage page,
        double centerX,
        double halfColumnWidth,
        double tableTop,
        double rangeTop)
    {
        var lines = page.RecognizedWords
            .Where(word =>
                word.CenterX >= centerX - halfColumnWidth &&
                word.CenterX <= centerX + halfColumnWidth &&
                word.CenterY > tableTop + 20 &&
                word.CenterY < rangeTop - 8)
            .GroupBy(word => word.LineIndex)
            .Select(ToLine)
            .Where(line =>
                !string.IsNullOrWhiteSpace(line.Text) &&
                !line.Text.StartsWith("高校", StringComparison.Ordinal))
            .OrderBy(line => line.Top)
            .Select(line => line.Text)
            .ToArray();

        return NormalizeMaterialName(string.Concat(lines));
    }

    private static string ExtractRange(
        RecognizedReportPage page,
        double centerX,
        double halfColumnWidth,
        double rangeCenterY)
    {
        var tolerance = page.PixelHeight * 0.018;
        return page.RecognizedWords
            .Where(word =>
                word.CenterX >= centerX - halfColumnWidth &&
                word.CenterX <= centerX + halfColumnWidth &&
                Math.Abs(word.CenterY - rangeCenterY) <= tolerance)
            .GroupBy(word => word.LineIndex)
            .Select(ToLine)
            .Where(line => RangeCandidateRegex().IsMatch(line.Text))
            .OrderBy(line => Math.Abs(line.CenterY - rangeCenterY))
            .Select(line => line.Text)
            .FirstOrDefault() ?? string.Empty;
    }

    private static OcrLine[] BuildLines(RecognizedReportPage page) =>
        page.RecognizedWords
            .GroupBy(word => word.LineIndex)
            .Select(ToLine)
            .OrderBy(line => line.Top)
            .ToArray();

    private static OcrLine ToLine(IEnumerable<RecognizedWord> words)
    {
        var ordered = words.OrderBy(word => word.Left).ToArray();
        var left = ordered.Min(word => word.Left);
        var top = ordered.Min(word => word.Top);
        var right = ordered.Max(word => word.Left + word.Width);
        var bottom = ordered.Max(word => word.Top + word.Height);
        return new OcrLine(
            string.Concat(ordered.Select(word => word.Text)).Replace(" ", string.Empty, StringComparison.Ordinal),
            left,
            top,
            right,
            bottom);
    }

    private static string NormalizeRange(string value)
    {
        var normalized = value.Replace('一', '-')
            .Replace('ー', '-')
            .Replace('〜', '-')
            .Replace('～', '-')
            .Replace('~', '-')
            .Replace(',', '-');
        return DigitSeparatedShinRegex().Replace(normalized, "-");
    }

    private static string NormalizeMaterialName(string value) =>
        value.Replace("ー新版】", "【新版】", StringComparison.Ordinal).Trim();

    private static string RepairRangeForMaterial(string range, string materialName) =>
        materialName.Contains("1900", StringComparison.Ordinal) &&
        string.Equals(range, "1-900", StringComparison.Ordinal)
            ? "1-1900"
            : range;

    [GeneratedRegex(@"^(\d{1,3})(?:問)?$")]
    private static partial Regex CountRegex();

    [GeneratedRegex(@"\d")]
    private static partial Regex RangeCandidateRegex();

    [GeneratedRegex(@"(?<=\d)新(?=\d)")]
    private static partial Regex DigitSeparatedShinRegex();

    private sealed record OcrLine(
        string Text,
        double Left,
        double Top,
        double Right,
        double Bottom)
    {
        public double CenterX => (Left + Right) / 2;

        public double CenterY => (Top + Bottom) / 2;
    }
}
