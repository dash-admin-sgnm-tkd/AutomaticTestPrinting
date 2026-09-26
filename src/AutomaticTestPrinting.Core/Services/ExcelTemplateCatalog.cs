using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static partial class ExcelTemplateCatalog
{
    public static ExcelRequestPreparation Prepare(
        NormalTestRequestCandidate request,
        string? materialFolder)
    {
        var catalog = ExcelTemplateProfileStore.LoadDefault();
        if (!catalog.IsValid)
        {
            return new ExcelRequestPreparation(
                false,
                false,
                $"教材設定を読み込めません。{catalog.Message}");
        }

        var normalizedMaterialName = NormalizeName(request.MaterialName);
        var profile = catalog.Profiles.FirstOrDefault(candidate =>
            candidate.MaterialNameKeywords.Any(keyword =>
                normalizedMaterialName.Contains(NormalizeName(keyword), StringComparison.Ordinal)));

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

        var availableCount = profile.RangeUsesSectionMapping
            ? profile.MaximumQuestionCount
            : end - start + 1;
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

        var workbookPath = FindWorkbook(materialFolder, profile);
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
            $"Excel連携準備OK：範囲 {start}-{end}／{request.QuestionCount}問／" +
            $"使用Excel：{Path.GetFileName(workbookPath)}",
            profile,
            workbookPath,
            start,
            end);
    }

    private static string? FindWorkbook(
        string? materialFolder,
        ExcelTemplateProfile profile)
    {
        if (string.IsNullOrWhiteSpace(materialFolder) || !Directory.Exists(materialFolder))
        {
            return null;
        }

        try
        {
            var normalizedKeyword = NormalizeName(profile.WorkbookNameKeyword);
            var excludedKeywords = profile.WorkbookNameExcludedKeywords
                .Select(NormalizeName)
                .Where(keyword => keyword.Length > 0)
                .ToArray();
            return Directory.EnumerateFiles(materialFolder, "*.xlsm", SearchOption.AllDirectories)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(path =>
                {
                    var normalizedFileName = NormalizeName(Path.GetFileNameWithoutExtension(path));
                    return normalizedFileName.Contains(normalizedKeyword, StringComparison.Ordinal) &&
                        !excludedKeywords.Any(excluded =>
                            normalizedFileName.Contains(excluded, StringComparison.Ordinal));
                });
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

    private static string NormalizeName(string value)
    {
        var normalized = value.Normalize(NormalizationForm.FormKC);
        return string.Concat(normalized
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant));
    }

    [GeneratedRegex(@"^\s*(\d+)\s*[-ー〜～~–—]\s*(\d+)\s*$")]
    private static partial Regex RangeRegex();
}
