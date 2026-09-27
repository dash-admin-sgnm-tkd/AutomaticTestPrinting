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
                    group.Select(profile => profile.Id).ToArray(),
                    CreatePreviewDescription(template),
                    CreatePreviewQuestion(template),
                    CreatePreviewAnswer(template));
            })
            .OrderBy(option => GetDisplayOrder(option.TemplateProfile))
            .ThenBy(option => option.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    private static MaterialFormatKey CreateFormatKey(ExcelTemplateProfile profile) => new(
        profile.FinishedLayoutId,
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
        if (!string.IsNullOrWhiteSpace(profile.FinishedLayoutId))
        {
            return profile.FinishedLayoutId;
        }

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
        if (!string.IsNullOrWhiteSpace(profile.FinishedLayoutName))
        {
            return profile.FinishedLayoutName;
        }

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

    private static string CreatePreviewDescription(ExcelTemplateProfile profile) =>
        CreateFormatId(profile) switch
        {
            "word-pair-list" => "英単語・英熟語を答える、短い1問1答形式です。",
            "grammar-choice" => "英文の空所や下線部を扱う、英文法・語法の問題形式です。",
            "japanese-vocabulary" => "国語の語彙と意味を対応させる1問1答形式です。",
            "balanced-two-column" => "問題と解答を同じくらいの幅で並べる形式です。",
            "question-wide-variable-row" => "長い問題文を広く表示し、文章量に応じて行の高さを変える形式です。",
            "answer-wide" => "短い問題に対して解答欄を広く取る形式です。",
            _ => "問題と解答を表にして印刷する形式です。"
        };

    private static string CreatePreviewQuestion(ExcelTemplateProfile profile) =>
        CreateFormatId(profile) switch
        {
            "word-pair-list" => "1. abandon\n2. accurate\n3. benefit",
            "grammar-choice" => "1. She (  ) to the library yesterday.\n   ① go  ② went  ③ gone  ④ going",
            "japanese-vocabulary" => "1. 『簡潔』の意味を答えなさい。",
            "question-wide-variable-row" => "1. 次の文章を読み、問いに答えなさい。\n   （長い問題文が入ります）",
            _ => "1. 問題文\n2. 問題文\n3. 問題文"
        };

    private static string CreatePreviewAnswer(ExcelTemplateProfile profile) =>
        CreateFormatId(profile) switch
        {
            "word-pair-list" => "1. 捨てる\n2. 正確な\n3. 利益",
            "grammar-choice" => "1. ② went",
            "japanese-vocabulary" => "1. 短く、要点がまとまっていること",
            "question-wide-variable-row" => "1. 解答例",
            _ => "1. 解答\n2. 解答\n3. 解答"
        };

    private sealed record MaterialFormatKey(
        string FinishedLayoutId,
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
