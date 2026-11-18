using System.Globalization;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class VerticalJapaneseQuestionLayout
{
    public const int MaximumColumns = 25;
    public const double TableWidth = 760;
    public const double QuestionHeight = 135;
    public const double AnswerHeight = 350;

    // Preserve existing Japanese registrations; no migration of user JSON is needed.
    public static bool UsesLayout(ExcelTemplateProfile profile) => StackedQuestionAnswerLayout.UsesLayout(profile);

    public static int ColumnsPerPage(IReadOnlyList<StackedQuestion> questions,
        double questionHeight = QuestionHeight, double answerHeight = AnswerHeight)
    {
        if (questions.Count == 0) throw new ArgumentException("問題がありません。", nameof(questions));
        if (questions.Select(q => Convert.ToString(q.Number, CultureInfo.InvariantCulture)).Distinct().Count() != questions.Count)
            throw new InvalidOperationException("選ばれた問題番号が重複しています。国語テストの作成を停止しました。");
        for (var columns = MaximumColumns; columns >= 1; columns--)
            if (questions.All(q => Fits(q.Question, questionHeight, TableWidth / columns, 8) &&
                Fits(q.Answer, answerHeight, TableWidth / columns, 8))) return columns;
        throw new InvalidOperationException("国語テストの文章が長すぎます。教材の問題・解答データを確認してください。");
    }

    public static double FontSize(string text, double height, double width)
    {
        for (double size = 10.5; size >= 8; size -= 0.5)
            if (Fits(text, height, width, size)) return size;
        throw new InvalidOperationException("縦書き欄に文章が収まりません。問題・解答データを確認してください。");
    }

    public static IReadOnlyList<object[,]> CreatePages(IReadOnlyList<StackedQuestion> questions, bool includeAnswers,
        double questionHeight = QuestionHeight, double answerHeight = AnswerHeight)
    {
        var columns = ColumnsPerPage(questions, questionHeight, answerHeight);
        var pages = new List<object[,]>();
        for (var start = 0; start < questions.Count; start += columns)
        {
            var count = Math.Min(columns, questions.Count - start);
            var values = new object[3, columns];
            for (var column = 0; column < columns; column++)
            {
                var offsetFromRight = columns - 1 - column;
                var question = offsetFromRight < count ? questions[start + offsetFromRight] : null;
                values[0, column] = question?.Number ?? "";
                values[1, column] = question?.Question ?? "";
                values[2, column] = includeAnswers ? question?.Answer ?? "" : "";
            }
            pages.Add(values);
        }
        return pages;
    }

    private static bool Fits(string text, double height, double width, double size)
    {
        // Estimate conservatively and keep explicit line breaks as separate vertical lines.
        var capacity = Math.Max(1, (int)((height - 6) / (size * 1.2)));
        var lines = text.Replace("\r", "", StringComparison.Ordinal).Split('\n')
            .Sum(line => Math.Max(1, (new StringInfo(line).LengthInTextElements + capacity - 1) / capacity));
        return lines * size + 6 <= width;
    }
}
