using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PolarisSubquestionExpanderTests
{
    [Fact]
    public void AggregateCorrectnessQuestionBecomesFourIndependentCorrections()
    {
        var source = new PolarisSourceQuestion(284, "3-8-C-1",
            "次の組には、文法または語法的に正しい英文はいくつありますか。①~④の中から1つ選びなさい。また、４つの英文がすべて誤りであると判断した場合は⑤を選びなさい。\n" +
            "(1) A fifteen-years-old boy won the speech contest.\n" +
            "(2) I apologized her for being late.\n" +
            "(3) Steve ran enough fast to catch the train.\n" +
            "(4) There used to be a bank at the corner.\n\n①1つ　②2つ　③3つ　④4つ　⑤すべて誤り",
            "①\n\n\n(1) years → year\n(2) her → to her\n(3) enough fast → fast enough\n(4) 誤りなし");
        var children = PolarisSubquestionExpander.Expand("final-polaris-1", source);

        Assert.Equal(4, children.Count);
        Assert.Equal(["years → year", "her → to her", "enough fast → fast enough", "誤りなし"], children.Select(c => c.Answer));
        Assert.Equal(["3-8-C-1(1)", "3-8-C-1(2)", "3-8-C-1(3)", "3-8-C-1(4)"], children.Select(c => c.Number));
        Assert.All(children, child => Assert.Equal(284, child.SourceNumber));
        Assert.All(children, child => Assert.DoesNotContain("①1つ", child.Question));
        Assert.DoesNotContain("apologized", children[0].Question);
        Assert.EndsWith("There used to be a bank at the corner.", children[3].Question);
        // Excel line ending differences do not turn an otherwise identical source into another edition.
        var windowsSource = source with { Question = source.Question.Replace("\n", "\r\n", StringComparison.Ordinal) };
        Assert.Equal(children, PolarisSubquestionExpander.Expand("final-polaris-1", windowsSource));
    }

    [Theory]
    [InlineData("final-polaris-1", 1, 286, 307)]
    [InlineData("final-polaris-1", 284, 284, 4)]
    [InlineData("final-polaris-1", 285, 286, 20)]
    [InlineData("final-polaris-1", 1, 283, 283)]
    [InlineData("final-polaris-2", 1, 305, 305)]
    [InlineData("final-polaris-3", 1, 273, 305)]
    [InlineData("final-polaris-3", 124, 130, 15)]
    [InlineData("final-polaris-3", 136, 138, 15)]
    [InlineData("final-polaris-3", 272, 273, 10)]
    [InlineData("vintage-4th", 1, 25, 25)]
    public void RangeCountsChildrenOfIncludedSourceNumbersOnly(string id, int start, int end, int expected) =>
        Assert.Equal(expected, PolarisSubquestionExpander.GetAvailableCount(id, start, end));

    [Theory]
    [InlineData("final-polaris-1", 283)]
    [InlineData("final-polaris-2", 45)]
    [InlineData("final-polaris-3", 44)]
    [InlineData("other", 284)]
    public void OrdinaryQuestionsAreNotSplitOnNumbersInTheirText(string id, int serial)
    {
        var source = new PolarisSourceQuestion(serial, "1-1", "Choose (1) or (2).", "(1)");
        Assert.Equal(source, Assert.Single(PolarisSubquestionExpander.Expand(id, source)));
    }

    [Fact]
    public void ChangedWorkbookStopsInsteadOfApplyingAnswersToAnotherEdition()
    {
        var source = new PolarisSourceQuestion(284, "3-8-C-1", "A replacement question", "A replacement answer");
        var error = Assert.Throws<InvalidOperationException>(() => PolarisSubquestionExpander.Expand("final-polaris-1", source));
        Assert.Contains("元番号 284", error.Message);
        Assert.Contains("この範囲を外して", error.Message);
    }
}
