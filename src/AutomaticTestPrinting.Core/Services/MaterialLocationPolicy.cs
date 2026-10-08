namespace AutomaticTestPrinting.Core.Services;

/// <summary>Choose subject-folder originals; never use the Kunitachi school copies.</summary>
public static class MaterialLocationPolicy
{
    private static readonly string[] SubjectFolders =
    ["①英語", "②数学", "③現代文", "④古文", "⑤漢文", "⑥物理", "⑦化学", "⑧生物",
     "⑨地学", "⑩日本史", "⑪世界史", "⑫地理", "⑬現代社会", "⑭倫理", "⑮情報", "⑯公共"];

    public static bool IsExcluded(string path) => DirectoryParts(path).Any(part =>
        part.StartsWith("国立校", StringComparison.Ordinal));

    public static bool IsPreferred(string path) => DirectoryParts(path).Any(part =>
        SubjectFolders.Any(subject => part.StartsWith(subject, StringComparison.Ordinal)) ||
        part.Contains("段階突破テスト", StringComparison.Ordinal));

    public static string[] SelectCandidates(IEnumerable<string> paths)
    {
        var eligible = paths.Where(path => !IsExcluded(path)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var preferred = eligible.Where(IsPreferred).ToArray();
        return (preferred.Length > 0 ? preferred : eligible).Order(StringComparer.OrdinalIgnoreCase).ToArray();
    }

    // Do not mistake a material filename for a containing directory.
    private static IEnumerable<string> DirectoryParts(string path) => path.Split(['/', '\\']).SkipLast(1);
}
