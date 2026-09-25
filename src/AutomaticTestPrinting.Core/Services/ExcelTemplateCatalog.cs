using System.Globalization;
using System.Text;
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
        "問題解答リスト",
        "講師用",
        "生徒用");

    private static readonly ExcelTemplateProfile EikenPre1Ex = new(
        "eiken-pre1-ex-second-edition",
        "英検準1級単熟語EX（第2版）",
        "準1級単熟語EX",
        ["英検準1級単熟語EX第2版", "英検準1級単熟語EX", "準1級単熟語EX"],
        2412,
        100,
        "作業シート",
        "G1",
        "G2",
        "問題解答リスト",
        "講師用",
        "生徒用");

    private static readonly ExcelTemplateProfile Target1000 = new(
        "target-1000-fifth-edition",
        "英熟語ターゲット1000（5訂版）",
        "英熟語ターゲット1000",
        ["英熟語ターゲット1000", "ターゲット1000"],
        1000,
        100,
        "作業シート",
        "G1",
        "G2",
        "問題解答リスト",
        "講師用",
        "生徒用");

    private static readonly ExcelTemplateProfile Vintage4 = new(
        "vintage-fourth-edition",
        "Vintage（4th Edition）",
        "Vintage",
        ["Vintage4thEdition", "Vintage"],
        1596,
        25,
        "作業シート",
        "G1",
        "G2",
        "問題解答リスト",
        "講師用",
        "生徒用");

    private static readonly ExcelTemplateProfile Vocabulary1700 = new(
        "vocabulary-1700-sigma-best",
        "国語力を伸ばす 語彙1700（シグマベスト）",
        "国語力を伸ばす 語彙1700",
        ["国語力を伸ばす語彙1700", "語彙1700"],
        1700,
        50,
        "作業シート",
        "G1",
        "G2",
        "問題解答リスト",
        "講師用",
        "生徒用",
        UseWideAnswerLayout: true);

    private static readonly ExcelTemplateProfile RapidReadingIdiomsRevised = new(
        "rapid-reading-idioms-revised",
        "速読英熟語（新版・改訂版）",
        "速読英熟語",
        ["速読英熟語改訂版", "速読英熟語"],
        74,
        50,
        "操作シート",
        "",
        "",
        "単語リスト",
        "講師用",
        "テスト用",
        UseWideAnswerLayout: true,
        RangeUsesSectionMapping: true,
        SectionMappingSheetName: "対応リスト",
        QuestionColumnWidth: 44,
        AnswerColumnWidth: 44,
        OutputRowHeight: 18);

    private static readonly IReadOnlyList<ExcelTemplateProfile> Profiles =
        [Target1900, EikenPre1Ex, Target1000, Vintage4, Vocabulary1700, RapidReadingIdiomsRevised];

    public static ExcelRequestPreparation Prepare(
        NormalTestRequestCandidate request,
        string? materialFolder)
    {
        var normalizedMaterialName = NormalizeName(request.MaterialName);
        var profile = Profiles.FirstOrDefault(candidate =>
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
            var normalizedKeyword = NormalizeName(fileNameKeyword);
            return Directory.EnumerateFiles(materialFolder, "*.xlsm", SearchOption.AllDirectories)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(path => NormalizeName(Path.GetFileNameWithoutExtension(path))
                    .Contains(normalizedKeyword, StringComparison.Ordinal));
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
