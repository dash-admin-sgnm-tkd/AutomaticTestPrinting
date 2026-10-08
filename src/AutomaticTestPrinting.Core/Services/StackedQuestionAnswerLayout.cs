using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public sealed record StackedQuestion(object Number, string Question, string Answer);

public static class StackedQuestionAnswerLayout
{
    public static bool UsesLayout(ExcelTemplateProfile profile) => profile.UseStackedQuestionAnswerLayout ||
        profile.FinishedLayoutId == "japanese-vocabulary";

    public static object[,] CreateValues(IReadOnlyList<StackedQuestion> questions, bool includeAnswers)
    {
        var values = new object[checked(questions.Count * 2), 3];
        for (var i = 0; i < questions.Count; i++)
        {
            values[i * 2, 0] = questions[i].Number;
            values[i * 2, 1] = questions[i].Question;
            values[i * 2, 2] = "";
            values[i * 2 + 1, 0] = "";
            values[i * 2 + 1, 1] = includeAnswers ? questions[i].Answer : "";
            values[i * 2 + 1, 2] = "";
        }
        return values;
    }
}
