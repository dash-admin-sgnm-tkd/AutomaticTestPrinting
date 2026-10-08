using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class VerticalJapaneseQuestionLayoutTests
{
    [Theory]
    [InlineData(1, 1)]
    [InlineData(25, 1)]
    [InlineData(26, 2)]
    [InlineData(50, 2)]
    [InlineData(100, 4)]
    public void PagesHaveHorizontalNumbersAndVerticalQuestionAnswerRows(int count, int pageCount)
    {
        var questions = Enumerable.Range(1, count).Select(i => new StackedQuestion(i, "語句", "意味です。")).ToArray();
        var teacher = VerticalJapaneseQuestionLayout.CreatePages(questions, true);
        var student = VerticalJapaneseQuestionLayout.CreatePages(questions, false);
        Assert.Equal(pageCount, teacher.Count);
        for (var i = 0; i < count; i++)
        {
            var page = i / 25;
            var column = i % 25;
            Assert.Equal(i + 1, teacher[page][0, column]);
            Assert.Equal("語句", teacher[page][1, column]);
            Assert.Equal("意味です。", teacher[page][2, column]);
            Assert.Equal("", student[page][2, column]);
        }
        if (count % 25 != 0) Assert.Equal("", teacher[^1][0, 24]);
    }

    [Fact]
    public void LongAnswersUseWiderColumnsWithoutDroppingQuestions()
    {
        var questions = Enumerable.Range(1, 25).Select(i => new StackedQuestion(i, "語彙", new string('あ', 180))).ToArray();
        var columns = VerticalJapaneseQuestionLayout.ColumnsPerPage(questions);
        Assert.InRange(columns, 1, 24);
        var pages = VerticalJapaneseQuestionLayout.CreatePages(questions, true);
        Assert.Equal(25, pages.Sum(page => Enumerable.Range(0, page.GetLength(1)).Count(c => page[0, c] is int)));
    }

    [Fact]
    public void ExistingJapaneseFlagsApplyButEnglishLayoutDoesNot()
    {
        Assert.True(VerticalJapaneseQuestionLayout.UsesLayout(new() { FinishedLayoutId = "japanese-vocabulary" }));
        Assert.True(VerticalJapaneseQuestionLayout.UsesLayout(new() { UseStackedQuestionAnswerLayout = true }));
        Assert.False(VerticalJapaneseQuestionLayout.UsesLayout(new() { FinishedLayoutId = "word-pair-list" }));
        Assert.Throws<InvalidOperationException>(() => VerticalJapaneseQuestionLayout.CreatePages(
            [new(1, "問1", "解1"), new(1, "問2", "解2")], true));
    }
}
