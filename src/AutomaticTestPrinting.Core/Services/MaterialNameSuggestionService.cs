using System.Text;

namespace AutomaticTestPrinting.Core.Services;

/// <summary>Suggestions only. Never select a book, edition, range or count automatically.</summary>
public static class MaterialNameSuggestionService
{
    public static IReadOnlyList<string> Suggest(string input, IEnumerable<(string Name, IEnumerable<string> Aliases)> materials)
    {
        var name = Normalize(input);
        if (name.Length < 5 || name.StartsWith("段階突破", StringComparison.Ordinal)) return [];
        var candidates = materials.Where(material => !Normalize(material.Name).StartsWith("段階突破", StringComparison.Ordinal))
            .Select(material => new { material.Name, Keys = material.Aliases.Append(material.Name).Select(Normalize).ToArray() }).ToArray();
        if (candidates.Any(candidate => candidate.Keys.Contains(name, StringComparer.Ordinal))) return [];
        return candidates.Select(candidate => new { candidate.Name, Score = candidate.Keys.Select(key => Similarity(name, key)).DefaultIfEmpty(0).Max() })
            .Where(candidate => candidate.Score >= 0.65).OrderByDescending(candidate => candidate.Score)
            .ThenBy(candidate => candidate.Name, StringComparer.Ordinal).Select(candidate => candidate.Name)
            .Distinct(StringComparer.Ordinal).Take(5).ToArray();
    }

    private static string Normalize(string value) => string.Concat(value.Normalize(NormalizationForm.FormKC)
        .Where(char.IsLetterOrDigit)).ToUpperInvariant();

    private static double Similarity(string left, string right)
    {
        if (right.Length < 5) return 0;
        var row = Enumerable.Range(0, right.Length + 1).ToArray();
        for (var i = 1; i <= left.Length; i++)
        {
            var diagonal = row[0];
            row[0] = i;
            for (var j = 1; j <= right.Length; j++)
            {
                var old = row[j];
                row[j] = Math.Min(Math.Min(row[j] + 1, row[j - 1] + 1), diagonal + (left[i - 1] == right[j - 1] ? 0 : 1));
                diagonal = old;
            }
        }
        return 1 - (double)row[right.Length] / Math.Max(left.Length, right.Length);
    }
}
