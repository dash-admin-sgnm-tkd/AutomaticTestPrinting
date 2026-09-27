using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ExcelMaterialFormatCatalog
{
    public static IReadOnlyList<ExcelMaterialFormatOption> CreateOptions(
        IEnumerable<ExcelTemplateProfile> profiles)
    {
        ArgumentNullException.ThrowIfNull(profiles);

        return profiles
            .GroupBy(CreateFormatKey)
            .Select(group =>
            {
                var template = group.First();
                return new ExcelMaterialFormatOption(
                    CreateFormatId(template),
                    CreateDisplayName(template),
                    template,
                    group.Select(profile => profile.Id).ToArray());
            })
            .OrderBy(option => GetDisplayOrder(option.TemplateProfile))
            .ThenBy(option => option.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    private static MaterialFormatKey CreateFormatKey(ExcelTemplateProfile profile) => new(
        profile.MaximumQuestionCount,
        profile.WorkingSheetName,
        profile.RangeStartCell,
        profile.RangeEndCell,
        profile.QuestionListSheetName,
        profile.TeacherSheetName,
        profile.StudentSheetName,
        profile.UseWideAnswerLayout,
        profile.RangeUsesSectionMapping,
        profile.SectionMappingSheetName,
        profile.SectionMappingUsesGridPairs,
        profile.SourceHasSeparateDisplayNumber,
        profile.QuestionColumnWidth,
        profile.AnswerColumnWidth,
        profile.OutputRowHeight,
        profile.AutoFitOutputRows,
        profile.TeacherHeaderText,
        profile.StudentHeaderText,
        profile.FitToSinglePageTall);

    private static string CreateFormatId(ExcelTemplateProfile profile)
    {
        if (profile.RangeUsesSectionMapping)
        {
            return profile.SectionMappingUsesGridPairs
                ? "chapter-grid"
                : "chapter-number";
        }

        if (profile.UseWideAnswerLayout)
        {
            return "wide-answer";
        }

        return profile.MaximumQuestionCount <= 25
            ? "grammar-25"
            : "vocabulary-standard";
    }

    private static string CreateDisplayName(ExcelTemplateProfile profile)
    {
        if (profile.RangeUsesSectionMapping)
        {
            return profile.SectionMappingUsesGridPairs
                ? "章・単元指定（公共・政治経済型）"
                : "章番号指定（速読英熟語型）";
        }

        if (profile.UseWideAnswerLayout)
        {
            return "国語・記述（横長解答型）";
        }

        return profile.MaximumQuestionCount <= 25
            ? "英文法・語法（25問型）"
            : "英単語・英熟語（標準型）";
    }

    private static int GetDisplayOrder(ExcelTemplateProfile profile)
    {
        if (!profile.RangeUsesSectionMapping &&
            !profile.UseWideAnswerLayout &&
            profile.MaximumQuestionCount > 25)
        {
            return 0;
        }

        if (!profile.RangeUsesSectionMapping &&
            !profile.UseWideAnswerLayout)
        {
            return 1;
        }

        if (!profile.RangeUsesSectionMapping)
        {
            return 2;
        }

        return profile.SectionMappingUsesGridPairs ? 4 : 3;
    }

    private sealed record MaterialFormatKey(
        int MaximumQuestionCount,
        string WorkingSheetName,
        string RangeStartCell,
        string RangeEndCell,
        string QuestionListSheetName,
        string TeacherSheetName,
        string StudentSheetName,
        bool UseWideAnswerLayout,
        bool RangeUsesSectionMapping,
        string? SectionMappingSheetName,
        bool SectionMappingUsesGridPairs,
        bool SourceHasSeparateDisplayNumber,
        double QuestionColumnWidth,
        double AnswerColumnWidth,
        double OutputRowHeight,
        bool AutoFitOutputRows,
        string TeacherHeaderText,
        string StudentHeaderText,
        bool FitToSinglePageTall);
}
