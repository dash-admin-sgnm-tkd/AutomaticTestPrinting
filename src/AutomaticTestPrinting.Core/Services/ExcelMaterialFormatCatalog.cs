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
        profile.UseWideAnswerLayout,
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
        if (!profile.UseWideAnswerLayout)
        {
            return "standard-table";
        }

        if (profile.AutoFitOutputRows)
        {
            return "question-wide-variable-row";
        }

        return Math.Abs(profile.QuestionColumnWidth - profile.AnswerColumnWidth) < 0.01
            ? "balanced-two-column"
            : profile.AnswerColumnWidth > profile.QuestionColumnWidth
                ? "answer-wide"
                : "question-wide";
    }

    private static string CreateDisplayName(ExcelTemplateProfile profile)
    {
        if (!profile.UseWideAnswerLayout)
        {
            return "標準（問題・解答の表）";
        }

        if (profile.AutoFitOutputRows)
        {
            return "問題欄ワイド（長文・自動行高）";
        }

        return Math.Abs(profile.QuestionColumnWidth - profile.AnswerColumnWidth) < 0.01
            ? "均等2列（問題・解答が同じ幅）"
            : profile.AnswerColumnWidth > profile.QuestionColumnWidth
                ? "解答欄ワイド（解答を広く表示）"
                : "問題欄ワイド（問題を広く表示）";
    }

    private static int GetDisplayOrder(ExcelTemplateProfile profile)
    {
        if (!profile.UseWideAnswerLayout)
        {
            return 0;
        }

        if (profile.AnswerColumnWidth > profile.QuestionColumnWidth)
        {
            return 1;
        }

        if (Math.Abs(profile.QuestionColumnWidth - profile.AnswerColumnWidth) < 0.01)
        {
            return 2;
        }

        return profile.AutoFitOutputRows ? 3 : 4;
    }

    private sealed record MaterialFormatKey(
        bool UseWideAnswerLayout,
        bool SourceHasSeparateDisplayNumber,
        double QuestionColumnWidth,
        double AnswerColumnWidth,
        double OutputRowHeight,
        bool AutoFitOutputRows,
        string TeacherHeaderText,
        string StudentHeaderText,
        bool FitToSinglePageTall);
}
