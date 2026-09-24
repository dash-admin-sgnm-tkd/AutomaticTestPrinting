using System.Globalization;
using System.Text.RegularExpressions;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static partial class ExcelTemplateCatalog
{
    private static readonly ExcelTemplateProfile Target1900 = new(
        "target-1900-sixth-edition",
        "ターゲット1900（6訂版）",
        "ターゲット1900",
        ["英単語ターゲット1900", "ターゲット1900"],
        1900,
        100,
        "作業シート",
        "G1",
        "G2",
        "講師用",
        "生徒用");

    private static readonly IReadOnlyList<ExcelTemplateProfile> Profiles = [Target1900];

    public static ExcelRequestPreparation Prepare(
        NormalTestRequestCandidate request,
        string? materialFolder)
    {
        var profile = Profiles.FirstOrDefault(candidate =>
            candidate.MaterialNameKeywords.Any(keyword =>
                request.MaterialName.Contains(keyword, StringComparison.OrdinalIgnoreCase)));

        if (profile is null)
        {
            return new ExcelRequestPreparation(
                false,
                false,
                "未対応の教材です。対応Excelの解析が必要です。");
        }

        var rangeMatch = RangeRegex().Match(request.Range);
        if (!rangeMatch.Success ||
            !int.TryParse(rangeMatch.Groups[1].Value, CultureInfo.InvariantCulture, out var start) ||
            !int.TryParse(rangeMatch.Groups[2].Value, CultureInfo.InvariantCulture, out var end))
        {
            return new ExcelRequestPreparation(
                true,
                false,
                "範囲を「開始番号-終了番号」の形で確認してください。",
                profile);
        }

        if (start < 1 || end < start || end > profile.MaximumQuestionNumber)
        {
            return new ExcelRequestPreparation(
                true,
                false,
                $"範囲は1-{profile.MaximumQuestionNumber}内で指定してください。",
                profile,
                StartNumber: start,
                EndNumber: end);
        }

        var availableCount = end - start + 1;
        if (request.QuestionCount < 1 ||
            request.QuestionCount > profile.MaximumQuestionCount ||
            request.QuestionCount > availableCount)
        {
            return new ExcelRequestPreparation(
                true,
                false,
                $"問題数は1-{Math.Min(profile.MaximumQuestionCount, availableCount)}問で指定してください。",
                profile,
                StartNumber: start,
                EndNumber: end);
        }

        var workbookPath = FindWorkbook(materialFolder, profile.WorkbookNameKeyword);
        if (workbookPath is null)
        {
            return new ExcelRequestPreparation(
                true,
                false,
                $"教材フォルダーに「{profile.WorkbookNameKeyword}」のExcelが見つかりません。",
                profile,
                StartNumber: start,
                EndNumber: end);
        }

        return new ExcelRequestPreparation(
            true,
            true,
            $"Excel連携準備OK：範囲 {start}-{end}／{request.QuestionCount}問",
            profile,
            workbookPath,
            start,
            end);
    }

    private static string? FindWorkbook(string? materialFolder, string fileNameKeyword)
    {
        if (string.IsNullOrWhiteSpace(materialFolder) || !Directory.Exists(materialFolder))
        {
            return null;
        }

        try
        {
            return Directory.EnumerateFiles(materialFolder, "*.xlsm", SearchOption.AllDirectories)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(path => Path.GetFileNameWithoutExtension(path)
                    .Contains(fileNameKeyword, StringComparison.OrdinalIgnoreCase));
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
        catch (IOException)
        {
            return null;
        }
    }

    [GeneratedRegex(@"^\s*(\d+)\s*[-ー〜～~–—]\s*(\d+)\s*$")]
    private static partial Regex RangeRegex();
}
