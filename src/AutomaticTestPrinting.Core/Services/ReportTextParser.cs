using System.Text.RegularExpressions;
using System.Globalization;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static partial class ReportTextParser
{
    public static RecognizedReport Parse(
        string sourcePath,
        IReadOnlyList<RecognizedReportPage> pages)
    {
        var text = string.Join(Environment.NewLine, pages.Select(page => page.Text));
        return new RecognizedReport(
            sourcePath,
            FindStudentName(text),
            FindNormalTestSummary(pages),
            FindStageTestSummary(text),
            pages);
    }

    private static string FindStudentName(string text)
    {
        var match = StudentNameRegex().Match(Compact(text));
        return match.Success ? match.Groups[1].Value.Trim() : "氏名を確認してください";
    }

    private static string FindNormalTestSummary(IReadOnlyList<RecognizedReportPage> pages)
    {
        var targetPage = pages.LastOrDefault(page =>
            Compact(page.Text).Contains("次回までの宿題", StringComparison.Ordinal));
        var compact = Compact(targetPage?.Text ?? string.Join(' ', pages.Select(page => page.Text)));
        if (!compact.Contains("テスト作成依頼", StringComparison.Ordinal))
        {
            return "通常テストの依頼欄を確認できませんでした";
        }

        var counts = QuestionCountRegex().Matches(compact)
            .Select(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture))
            .Where(count => count is > 0 and <= 200)
            .ToArray();

        return counts.Length == 0
            ? "依頼欄あり（問題数は画面で確認してください）"
            : $"問題数のOCR候補：{string.Join("・", counts.Select(count => $"{count}問"))}（原本確認）";
    }

    private static string FindStageTestSummary(string text)
    {
        var compact = Compact(text);
        var marker = compact.LastIndexOf("次回の内容", StringComparison.Ordinal);
        if (marker < 0)
        {
            return compact.Contains("段階突破", StringComparison.Ordinal)
                ? "段階突破の記載あり（内容を確認してください）"
                : "段階突破テストの依頼なし";
        }

        var section = compact[marker..];
        var subject = StageSubjectRegex().Match(section);
        var details = StageDetailsRegex().Matches(section).LastOrDefault();
        if (!subject.Success || details is null)
        {
            return "段階突破テストの依頼欄あり（内容を確認してください）";
        }

        return $"段階突破：{subject.Groups[1].Value}・{details.Groups[1].Value}年度・第{details.Groups[3].Value}回";
    }

    private static string Compact(string value) =>
        WhitespaceRegex().Replace(value.Replace('：', ':').Replace('．', '.'), string.Empty);

    [GeneratedRegex(@"生徒名[:.]?(.+?)(?=(?:教務チェック|講師名|20\d{2}年|通常特訓|科目名|$))")]
    private static partial Regex StudentNameRegex();

    [GeneratedRegex(@"(?<!\d)(\d{1,3})\s*問")]
    private static partial Regex QuestionCountRegex();

    [GeneratedRegex(@"科目(.+?)(?=(?:得点|年度))")]
    private static partial Regex StageSubjectRegex();

    [GeneratedRegex(@"(?:年度)?(20\d{2})(?:年度)?段階(.+?)回(\d{1,2})")]
    private static partial Regex StageDetailsRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();
}
