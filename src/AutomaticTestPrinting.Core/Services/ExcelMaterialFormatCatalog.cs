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
        StackedQuestionAnswerLayout.UsesLayout(profile),
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
        if (StackedQuestionAnswerLayout.UsesLayout(profile))
        {
            return "国語・漢字（縦書き・問題と解答欄を上下）";
        }
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
            "grammar-choice" => "問題文と選択肢を広く表示する、文法・選択問題向けの形式です。",
            "reading-landscape" => "A4横向き・1ページ5問。英文と解答欄を左右に配置します。解答用は和訳・構文解説の欄を広くし、長い文章は文字の大きさを調整します。",
            "reading-portrait" => "A4縦向き・1ページ10問。英文と和訳・構文解説を左右に配置し、問題と解答の文字サイズを別々に調整します。",
            "japanese-vocabulary" => "A4横向き。各列の上に番号、中央に縦書きの問題、下に解答欄を配置します。標準25問／ページ、長い文章は列数を減らして改ページします。",
            "balanced-two-column" => "問題と解答を同じくらいの幅で並べる形式です。",
            "question-wide-variable-row" => "長い問題文を広く表示し、文章量に応じて行の高さを変える形式です。",
            "answer-wide" => "短い問題に対して解答欄を広く取る形式です。",
            _ when profile.UseStandardQuestionTable => $"A4縦向き・1ページ{profile.OriginalGrammarQuestionsPerPage}問。問題と解答の文字サイズを別々に調整し、下線・太字・上付き・下付き文字を引き継ぎます。" +
                (profile.SectionMappingUsesRowPairs ? "範囲は単元番号で指定します。" : profile.SourceRangeColumn.Length>0 ? "範囲はテーマ番号で指定します。" : ""),
            _ => "問題と解答を表にして印刷する形式です。"
        };

    private static string CreatePreviewQuestion(ExcelTemplateProfile profile) =>
        CreateFormatId(profile) switch
        {
            "word-pair-list" => "1. abandon\n2. accurate\n3. benefit",
            "grammar-choice" => "1. 問題文（空所・下線部など）\n   ① 選択肢  ② 選択肢  ③ 選択肢",
            "reading-landscape" => "英文（左）　｜　和訳・構文の記入欄（右）\n1ページ5問で改ページ",
            "reading-portrait" => "英文（左）　｜　和訳・構文の記入欄（右）\n1ページ10問で改ページ",
            "japanese-vocabulary" => "上：問題番号\n中央：語句・読み・書きを縦書き表示\n下：解答欄（生徒用は空欄）",
            "question-wide-variable-row" => "1. 次の文章を読み、問いに答えなさい。\n   （長い問題文が入ります）",
            _ => "1. 問題文\n2. 問題文\n3. 問題文"
        };

    private static string CreatePreviewAnswer(ExcelTemplateProfile profile) =>
        CreateFormatId(profile) switch
        {
            "word-pair-list" => "1. 捨てる\n2. 正確な\n3. 利益",
            "grammar-choice" => "1. ② 正答",
            "reading-landscape" => "和訳・構文解説（生徒用は空欄）",
            "reading-portrait" => "和訳・構文解説（生徒用は空欄）",
            "japanese-vocabulary" => "1. 解答（意味・読み・漢字など）",
            "question-wide-variable-row" => "1. 解答例",
            _ => "1. 解答\n2. 解答\n3. 解答"
        };

    private sealed record MaterialFormatKey(
        string FinishedLayoutId,
        bool UseStackedQuestionAnswerLayout,
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
