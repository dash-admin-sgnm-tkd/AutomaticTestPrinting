using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public class StackedQuestionAnswerLayoutTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CreateValues_PlacesEachQuestionAboveItsAnswer(bool includeAnswers)
    {
        var questions = Enumerable.Range(1, 100)
            .Select(i => new StackedQuestion(i, $"問題{i}", $"解答{i}")).ToArray();
        var values = StackedQuestionAnswerLayout.CreateValues(questions, includeAnswers);
        Assert.Equal(200, values.GetLength(0));
        for (var i = 0; i < 100; i++)
        {
            Assert.Equal(i + 1, values[i * 2, 0]);
            Assert.Equal($"問題{i + 1}", values[i * 2, 1]);
            Assert.Equal(includeAnswers ? $"解答{i + 1}" : "", values[i * 2 + 1, 1]);
            Assert.Equal("", values[i * 2 + 1, 0]);
        }
    }

    [Fact]
    public void UsesLayout_AppliesToExistingJapaneseRegistrations()
    {
        Assert.True(StackedQuestionAnswerLayout.UsesLayout(new ExcelTemplateProfile
        {
            FinishedLayoutId = "japanese-vocabulary"
        }));
        Assert.False(StackedQuestionAnswerLayout.UsesLayout(new ExcelTemplateProfile()));
    }
}
