using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static partial class ExcelTemplateCatalog
{
    public static ExcelRequestPreparation Prepare(
        NormalTestRequestCandidate request,
        string? materialFolder) =>
        Prepare(
            request,
            string.IsNullOrWhiteSpace(materialFolder) ? [] : [materialFolder]);

    public static ExcelRequestPreparation Prepare(
        NormalTestRequestCandidate request,
        IReadOnlyCollection<string> materialFolders,
        IReadOnlyList<string>? knownWorkbookCandidates = null,
        string? workbookChoicesPath = null,
        IReadOnlyList<ExcelTemplateProfile>? profiles = null)
    {
        var catalog = profiles is null ? ExcelTemplateProfileStore.LoadDefault() :
            new ExcelTemplateProfileLoadResult(true, "", profiles, "明示的な教材設定");
        if (!catalog.IsValid)
        {
            return new ExcelRequestPreparation(
                false,
                false,
                $"教材設定を読み込めません。{catalog.Message}");
        }

        var normalizedMaterialName = NormalizeName(request.MaterialName);
        var (mappedProfile, aliasReview) = ReportExcelAliases.Resolve(catalog.Profiles, request.MaterialName);
        var matches = mappedProfile is null ? ExcelMaterialNameMatcher.FindBestMatches(catalog.Profiles, request.MaterialName) : [mappedProfile];
        if (matches.Count > 1)
            return new ExcelRequestPreparation(true, false,
                "教材名が複数の登録に一致しています。カタログで教材名・略称を区別してください：" +
                string.Join(" / ", matches.Select(p => p.DisplayName)));
        var profile = matches.Count == 1 ? matches[0] : null;

        if (profile is null)
        {
            if (aliasReview is not null)
                return new ExcelRequestPreparation(true, false, aliasReview);

            try
            {
                if (SharedMaterialCatalog.Load().Items.Any(item => item.Kind == "Excel" &&
                    item.Aliases.Append(item.DisplayName).Any(alias => NormalizeName(alias) == normalizedMaterialName)))
                    return new ExcelRequestPreparation(false, false,
                        "共通カタログにあるExcel教材ですが作成設定は未登録です。「教材カタログ」から使用ファイル・形式・上限を確認して登録してください。");
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ArgumentException)
            {
                return new ExcelRequestPreparation(false, false, $"共通教材カタログを確認してください：{exception.Message}");
            }
            return new ExcelRequestPreparation(
                false,
                false,
                "教材名がExcel・PDFの登録と一致しません。読み取り結果・版・略称を原本と照合し、未登録なら「教材カタログ」または「PDF教材を登録」で設定してください。");
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

        if (start < profile.MinimumQuestionNumber || end < start || end > profile.MaximumQuestionNumber)
        {
            return new ExcelRequestPreparation(
                true,
                false,
                $"範囲は{profile.MinimumQuestionNumber}-{profile.MaximumQuestionNumber}内で指定してください。",
                profile,
                StartNumber: start,
                EndNumber: end);
        }

        var availableCount = UnitQuestionRange.CandidateCount(profile,start,end);
        if (availableCount <= 0)
            return new ExcelRequestPreparation(true, false, "指定範囲には有効な問題がありません。開始・終了番号を確認してください。", profile,
                StartNumber: start, EndNumber: end);
        if (request.QuestionCount < 1 ||
            request.QuestionCount > profile.EffectiveMaximumQuestionCount ||
            request.QuestionCount > availableCount)
        {
            return new ExcelRequestPreparation(
                true,
                false,
                $"問題数は1-{Math.Min(profile.EffectiveMaximumQuestionCount, availableCount)}問で指定してください。",
                profile,
                StartNumber: start,
                EndNumber: end);
        }

        WorkbookResolution resolution;
        try
        {
            WorkbookChoiceStore.Load(workbookChoicesPath).TryGetValue(profile.Id, out var choice);
            resolution = WorkbookResolutionService.Resolve(materialFolders, profile, choice, knownWorkbookCandidates);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            return new ExcelRequestPreparation(true, false, exception.Message, profile, StartNumber: start, EndNumber: end);
        }
        var workbookPath = resolution.Path;
        if (workbookPath is null && resolution.Message.Length > 0)
            return new ExcelRequestPreparation(true, false, resolution.Message, profile, StartNumber: start,
                EndNumber: end, WorkbookCandidates: resolution.Candidates);
        if (workbookPath is null)
        {
            return new ExcelRequestPreparation(
                true,
                false,
                $"登録したExcel「{(string.IsNullOrEmpty(profile.WorkbookFileName) ? profile.WorkbookNameKeyword + ".xlsm" : profile.WorkbookFileName)}」が見つかりません。まず教材フォルダーの指定と同期を確認してください。改名や差し替えの場合は「登録Excelを変更」で使うファイルを選び直してください。処理を続けるには、この依頼のチェックを外せます。",
                profile,
                StartNumber: start,
                EndNumber: end);
        }

        return new ExcelRequestPreparation(
            true,
            true,
            $"Excel連携準備OK：範囲 {start}-{end}／{request.QuestionCount}問／" +
            $"使用Excel：{workbookPath}" +
            (PolarisSubquestionExpander.Supports(profile)
                ? $"\n元の連番 {start}-{end} を小問へ分割して抽選します（候補 {availableCount}問）。" : "") +
            (profile.GroupDuplicateQuestionNumbers ? $"\n元の単語番号で範囲を指定し、派生語の候補 {availableCount}問から抽選します。" : "") +
            (profile.SectionMappingUsesRowPairs ? "\nこの教材は単元番号で範囲を指定します。対応表から単語番号へ変換します。" : "") +
            (profile.SourceRangeColumn.Length>0 ? "\nこの教材はテーマ番号で範囲を指定します。各テーマに含まれる問題から抽選します。" : "") +
            (resolution.Message.Length > 0 ? $"\n{resolution.Message}" : ""),
            profile,
            workbookPath,
            start,
            end,
            resolution.Candidates);
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
