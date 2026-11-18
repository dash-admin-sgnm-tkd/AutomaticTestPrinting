using System.Text.RegularExpressions;

namespace AutomaticTestPrinting.Core.Services;

public static partial class ReadingTextFormatter
{
    public static string Normalize(string text, bool reduceInlinePadding)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n')
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => reduceInlinePadding ? PaddingRegex().Replace(line.TrimEnd(), " ") : line.TrimEnd());
        return string.Join('\n', lines);
    }

    [GeneratedRegex("[ \\t　]{4,}")]
    private static partial Regex PaddingRegex();
}
