using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class StagePdfIdentityTests
{
    [Theory]
    [InlineData("段階突破テストver.2020/日大レベル/古文/古文日大レベルver.2020-2_20210208本部.pdf", "古文", "2020", "日大", "2")]
    [InlineData("現代文日大レベルver.2020-2_20210208本部.pdf", "現代文", "2020", "日大", "2")]
    [InlineData("英語(文法)MARCHレベルver.2021-1.pdf", "英語(文法)", "2021", "MARCH", "1")]
    [InlineData("現代文日大レベルver2018-2_20190914本部.pdf", "現代文", "2018", "日大", "2")]
    [InlineData("政治・経済_地方国公立レベル2018-1.pdf", "政治・経済", "2018", "地方国公立", "1")]
    public void Parse_KeepsEveryIdentityDimension(string path, string subject, string year, string level, string attempt)
    {
        Assert.Equal(new StagePdfIdentity(subject, year, level, attempt), StagePdfIdentityParser.Parse(path));
    }

    [Theory]
    [InlineData("2024年英語段階突破テストvol.1.pdf")]
    [InlineData("英語vol1 配点表.pdf")]
    [InlineData("学研/古文_日大レベル2018-1.pdf")]
    [InlineData("古文日大レベルver.2020-2_解答.pdf")]
    [InlineData("古文日大レベルver.2020-2 - コピー.pdf")]
    public void Parse_DoesNotInferMissingOrAlternativeIdentity(string path) => Assert.Null(StagePdfIdentityParser.Parse(path));
}
