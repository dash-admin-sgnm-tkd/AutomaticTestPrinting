using System.Text;
using System.Text.RegularExpressions;

namespace AutomaticTestPrinting.Core.Services;

public sealed record StagePdfIdentity(string Subject, string Year, string Level, string Attempt)
{
    public string DisplayName => $"段階突破 {Subject} {Year}年度 {Level} 第{Attempt}回";
    public string LevelAlias => $"段階突破 {Subject} {Year}年度 {Level}レベル 第{Attempt}回";
}

public static partial class StagePdfIdentityParser
{
    public static StagePdfIdentity? Parse(string relativePath)
    {
        // Supplementary score sheets, alternate publisher versions, and year-only tests are not inferred.
        var normalized = relativePath.Normalize(NormalizationForm.FormKC);
        if (normalized.Contains("学研", StringComparison.Ordinal) || normalized.Contains("配点", StringComparison.Ordinal) ||
            normalized.Contains("解答", StringComparison.Ordinal) || normalized.Contains("問題のみ", StringComparison.Ordinal)) return null;
        var stem = string.Concat(Path.GetFileNameWithoutExtension(normalized.Replace('\\', '/')).Where(c => !char.IsWhiteSpace(c)));
        var match = FileIdentityRegex().Match(stem);
        if (!match.Success) return null;
        return new(match.Groups["subject"].Value, match.Groups["year"].Value,
            match.Groups["level"].Value, match.Groups["attempt"].Value);
    }

    [GeneratedRegex(@"^(?:★)?(?<subject>現代文|古文|漢文|英語(?:\((?:長文|文法)\))?|文系数学|理系数学|世界史|日本史|政治・?経済|地理|化学|物理|生物)_?(?<level>日大|MARCH|地方国公立|難関国公立|早慶|早稲田|東大)レベル(?:ver\.?)?(?<year>20\d{2})-(?<attempt>[12])\.?(?:_\d{8}(?:本部)?)?$")]
    private static partial Regex FileIdentityRegex();
}
