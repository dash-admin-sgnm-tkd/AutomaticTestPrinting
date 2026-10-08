using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ExcelMaterialNameMatcher
{
    public static IReadOnlyList<ExcelTemplateProfile> FindBestMatches(
        IReadOnlyList<ExcelTemplateProfile> profiles, string materialName)
    {
        var name = PdfMaterialProfileStore.NormalizeName(materialName);
        if (name.Length == 0) return [];
        var ranked = profiles.Select(profile => new
        {
            Profile = profile,
            Score = profile.MaterialNameKeywords.Append(profile.DisplayName)
                .Select(PdfMaterialProfileStore.NormalizeName)
                .Where(key => key.Length > 0 && name.Contains(key, StringComparison.Ordinal))
                .Select(key => key == name ? int.MaxValue : key.Length).DefaultIfEmpty(0).Max()
        }).Where(match => match.Score > 0).ToArray();
        if (ranked.Length == 0) return [];
        var best = ranked.Max(match => match.Score);
        return ranked.Where(match => match.Score == best).Select(match => match.Profile).ToArray();
    }
}
