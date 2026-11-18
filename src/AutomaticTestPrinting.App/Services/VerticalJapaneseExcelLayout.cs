using System.Runtime.InteropServices;
using System.Globalization;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.App.Services;

public static class VerticalJapaneseExcelLayout
{
    public static void Apply(dynamic sheet, IReadOnlyList<StackedQuestion> questions,
        string title, string range, bool includeAnswers,
        IReadOnlyDictionary<string, IReadOnlyList<(int Start, int Length)>>? questionEmphasis = null,
        bool balancedKanjiRows = false)
    {
        var questionHeight = balancedKanjiRows ? 242.5 : questionEmphasis is not null ? 330 : VerticalJapaneseQuestionLayout.QuestionHeight;
        var answerHeight = balancedKanjiRows ? 242.5 : questionEmphasis is not null ? 155 : VerticalJapaneseQuestionLayout.AnswerHeight;
        var columns = VerticalJapaneseQuestionLayout.ColumnsPerPage(questions, questionHeight, answerHeight);
        var pages = VerticalJapaneseQuestionLayout.CreatePages(questions, includeAnswers, questionHeight, answerHeight);
        var endColumn = (char)('B' + columns - 1);
        var lastRow = pages.Count * 4 - 1;
        object? used = null, output = null, font = null, setup = null, columnRange = null;
        try
        {
            used = sheet.UsedRange;
            ((dynamic)used).UnMerge();
            ((dynamic)used).Clear();
            output = sheet.Range[$"B1:{endColumn}{lastRow}"];
            ((dynamic)output).NumberFormat = "@";
            ((dynamic)output).WrapText = true;
            ((dynamic)output).HorizontalAlignment = -4108;
            ((dynamic)output).VerticalAlignment = -4108;
            font = ((dynamic)output).Font;
            ((dynamic)font).Name = "ＭＳ 明朝";
            ((dynamic)font).Size = 10.5;
            ((dynamic)font).Color = 0;
            columnRange = sheet.Columns[$"B:{endColumn}"];
            ((dynamic)columnRange).ColumnWidth = 4.4;
            for (var i = 0; i < 3; i++)
                ((dynamic)columnRange).ColumnWidth = (double)((dynamic)columnRange).ColumnWidth *
                    VerticalJapaneseQuestionLayout.TableWidth / (double)((dynamic)output).Width;
            sheet.ResetAllPageBreaks();
            setup = sheet.PageSetup;
            ((dynamic)setup).Orientation = 2;
            ((dynamic)setup).PaperSize = 9;
            ((dynamic)setup).Zoom = false;
            ((dynamic)setup).FitToPagesWide = 1;
            ((dynamic)setup).FitToPagesTall = false;
            ((dynamic)setup).CenterHorizontally = true;
            ((dynamic)setup).CenterVertically = false;
            ((dynamic)setup).TopMargin = 40;
            ((dynamic)setup).BottomMargin = 28;
            ((dynamic)setup).LeftMargin = 30;
            ((dynamic)setup).RightMargin = 30;
            ((dynamic)setup).HeaderMargin = 16;
            ((dynamic)setup).FooterMargin = 12;
            ((dynamic)setup).PrintTitleRows = "";
            ((dynamic)setup).PrintTitleColumns = "";
            ((dynamic)setup).PrintGridlines = false;
            ((dynamic)setup).PrintHeadings = false;
            ((dynamic)setup).PrintArea = $"B1:{endColumn}{lastRow}";
            ((dynamic)setup).LeftHeader = "&10 " + EscapeHeader(title);
            ((dynamic)setup).CenterHeader = includeAnswers ? "&10 講師用" : "&10 生徒用";
            ((dynamic)setup).RightHeader = $"&10 範囲：{EscapeHeader(range)} 問題数：{questions.Count}問";
            ((dynamic)setup).LeftFooter = "";
            ((dynamic)setup).CenterFooter = pages.Count > 1 ? "&8 &P / &N" : "";
            ((dynamic)setup).RightFooter = "";

            for (var page = 0; page < pages.Count; page++)
            {
                var top = page * 4 + 1;
                WithRange((object)sheet, $"B{top}:{endColumn}{top + 2}", block =>
                {
                    block.Value2 = pages[page];
                    object? borders = null;
                    try { borders = block.Borders; ((dynamic)borders).LineStyle = 1; ((dynamic)borders).Weight = 2; }
                    finally { Release(borders); }
                });
                WithRange((object)sheet, $"B{top}:{endColumn}{top}", numbers => { numbers.Orientation = 0; numbers.RowHeight = 16; });
                WithRange((object)sheet, $"B{top + 1}:{endColumn}{top + 1}", body =>
                { body.Orientation = -4166; body.RowHeight = questionHeight; });
                WithRange((object)sheet, $"B{top + 2}:{endColumn}{top + 2}", answers =>
                { answers.Orientation = -4166; answers.RowHeight = answerHeight; });
                for (var column = 0; column < columns; column++)
                {
                    var index = page * columns + column;
                    if (index >= questions.Count) break;
                    var question = questions[index].Question;
                    var emphasis = questionEmphasis?.GetValueOrDefault(question);
                    var addressColumn = (char)(endColumn - column);
                    SetFont(sheet, $"{addressColumn}{top + 1}", question,
                        questionHeight, false, emphasis);
                    SetFont(sheet, $"{addressColumn}{top + 2}", questions[index].Answer,
                        answerHeight, includeAnswers);
                }
                if (!includeAnswers)
                    WithRange((object)sheet, $"B{top + 2}:{endColumn}{top + 2}", answers =>
                    {
                        object? values = answers.Value2;
                        if (values is object[,] cells)
                        {
                            foreach (var value in cells)
                                if (!string.IsNullOrEmpty(Convert.ToString(value, CultureInfo.InvariantCulture)))
                                    throw new InvalidOperationException("問題PDFの解答欄に値が残っています。");
                        }
                        else if (!string.IsNullOrEmpty(Convert.ToString(values, CultureInfo.InvariantCulture)))
                            throw new InvalidOperationException("問題PDFの解答欄に値が残っています。");
                    });
                if (page > 0)
                {
                    object? anchor = null, breaks = null, added = null;
                    try { anchor = sheet.Range[$"B{top}"]; breaks = sheet.HPageBreaks; added = ((dynamic)breaks).Add(anchor); }
                    finally { Release(added); Release(breaks); Release(anchor); }
                }
                if (page + 1 < pages.Count)
                    WithRange((object)sheet, $"B{top + 3}:{endColumn}{top + 3}", spacer => spacer.RowHeight = 8);
            }
        }
        finally { Release(setup); Release(columnRange); Release(font); Release(output); Release(used); }
    }

    private static string EscapeHeader(string text) => text.Replace("&", "&&", StringComparison.Ordinal);

    private static void SetFont(dynamic sheet, string address, string text, double height, bool blue,
        IReadOnlyList<(int Start, int Length)>? emphasis = null)
    {
        WithRange((object)sheet, address, cell =>
        {
            object? font = null;
            try
            {
                font = cell.Font;
                ((dynamic)font).Size = VerticalJapaneseQuestionLayout.FontSize(text, height, (double)cell.Width);
                ((dynamic)font).Color = blue ? 0xA15D17 : 0;
                if (emphasis is not null)
                {
                    ((dynamic)font).Bold = false;
                    foreach (var (start, length) in emphasis)
                    {
                        object? characters = null, characterFont = null;
                        try
                        {
                            characters = cell.Characters[start, length];
                            characterFont = ((dynamic)characters).Font;
                            ((dynamic)characterFont).Bold = true;
                        }
                        finally { Release(characterFont); Release(characters); }
                    }
                }
            }
            finally { Release(font); }
        });
    }

    private static void WithRange(dynamic sheet, string address, Action<dynamic> action)
    {
        object? range = null;
        try { range = sheet.Range[address]; action((dynamic)range); }
        finally { Release(range); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
