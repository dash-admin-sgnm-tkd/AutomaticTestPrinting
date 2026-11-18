using System.Globalization;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public sealed record ConfiguredQuestion(int SourceNumber, string Number, string Question, string Answer, int SourceRow = 0);

public static class ConfiguredQuestionTableReader
{
    public static IReadOnlyList<ConfiguredQuestion> Read(object[,] values, ExcelTemplateProfile profile, int start, int end)
    {
        var questions = new Dictionary<int, List<ConfiguredQuestion>>();
        var firstRow = values.GetLowerBound(0);
        if (!int.TryParse(Text(firstRow, profile.SourceNumberColumn), NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
        {
            var questionHeader = Text(firstRow, profile.SourceQuestionColumn).Trim();
            var answerHeader = Text(firstRow, profile.SourceAnswerColumn).Trim();
            if (new[] { "番号", "連番", "難易度", "テーマ", "テーマ番号", "解答", "訳", "RND", "乱数" }.Contains(questionHeader, StringComparer.OrdinalIgnoreCase) ||
                new[] { "番号", "連番", "RND", "乱数" }.Contains(answerHeader, StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException("元Excelの列見出しと問題・解答の設定が一致しません。教材の形式を選び直してください。");
        }
        for (var row = values.GetLowerBound(0); row <= values.GetUpperBound(0); row++)
        {
            if (!int.TryParse(Text(row, profile.SourceNumberColumn), NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
                continue;
            var rangeNumber = number;
            if (profile.SourceRangeColumn.Length > 0)
            {
                var rangeText=Text(row,profile.SourceRangeColumn);
                if (profile.SourceRangeUsesPrefix && rangeText.IndexOf('-') is var separator && separator>=0) rangeText=rangeText[..separator];
                if (!int.TryParse(rangeText,NumberStyles.Integer,CultureInfo.InvariantCulture,out rangeNumber)) continue;
            }
            if (rangeNumber < start || rangeNumber > end) continue;
            var question = Text(row, profile.SourceQuestionColumn);
            var answer = Text(row, profile.SourceAnswerColumn);
            if (string.IsNullOrWhiteSpace(question) || string.IsNullOrWhiteSpace(answer)) continue;
            if (question.StartsWith("#REF!", StringComparison.Ordinal) || answer.StartsWith("#REF!", StringComparison.Ordinal))
                throw new InvalidOperationException($"問題番号 {number} の問題・解答にExcel計算エラーがあります。元Excelを確認してください。");
            var display = profile.SourceDisplayNumberColumn.Length == 0 ? "" : Text(row, profile.SourceDisplayNumberColumn);
            var data = new ConfiguredQuestion(number, string.IsNullOrWhiteSpace(display) ? number.ToString(CultureInfo.InvariantCulture) : display, question, answer,
                row - values.GetLowerBound(0) + 1);
            if (!questions.TryGetValue(number, out var group)) questions[number] = group = [];
            if (group.Any(existing => existing.Number == data.Number && existing.Question == data.Question && existing.Answer == data.Answer)) continue;
            if (group.Count > 0 && !profile.GroupDuplicateQuestionNumbers)
                throw new InvalidOperationException($"問題番号 {number} に異なる問題・解答が重複しています。教材の版・列設定を確認してください。");
            group.Add(data);
        }
        return questions.Values.SelectMany(group => group.Select((question, index) => group.Count > 1
            ? question with { Number = $"{question.Number}({index + 1})" } : question)).ToArray();

        string Text(int row, string column)
        {
            if (column.Length != 1 || column[0] < 'A' || column[0] > 'H')
                throw new ArgumentException("元データの列はA〜Hで指定してください。", nameof(profile));
            var index = values.GetLowerBound(1) + column[0] - 'A';
            return index <= values.GetUpperBound(1) ? Convert.ToString(values[row, index], CultureInfo.InvariantCulture) ?? "" : "";
        }
    }
}
