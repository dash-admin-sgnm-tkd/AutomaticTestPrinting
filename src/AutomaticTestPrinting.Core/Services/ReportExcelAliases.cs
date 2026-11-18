using System.Text.Json;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ReportExcelAliases
{
    private static readonly string[] VariantQualifiers = ["派生語", "多義語", "穴埋め", "英単語を解答", "熟語編"];
    private sealed record Entry(string Name, string[] Files);

    private static readonly Lazy<Dictionary<string, Entry>> Entries = new(() =>
    {
        using var stream = typeof(ReportExcelAliases).Assembly.GetManifestResourceStream(
            "AutomaticTestPrinting.Core.Configuration.report-excel-aliases.json")
            ?? throw new InvalidOperationException("レポート表記名の対応表を読み込めません。");
        return (JsonSerializer.Deserialize<Entry[]>(stream) ?? [])
            .ToDictionary(entry => PdfMaterialProfileStore.NormalizeName(entry.Name), StringComparer.Ordinal);
    });

    public static (ExcelTemplateProfile? Profile, string? ReviewMessage) Resolve(
        IReadOnlyList<ExcelTemplateProfile> profiles, string name)
    {
        if (!Entries.Value.TryGetValue(PdfMaterialProfileStore.NormalizeName(name), out var entry))
            return (null, null);

        if (entry.Files.Length != 1)
            return (null, "同じレポート表記名に複数のExcel候補があります。版・用途を確認してください：" +
                string.Join(" ／ ", entry.Files.Select(file => file.Length == 0 ? "ファイル名未記入" : file)));

        var file = entry.Files[0];
        var matches = file.Length == 0 ? [] : profiles.Where(profile => profile.WorkbookFileName.Length > 0 &&
            string.Equals(WithoutXlsm(profile.WorkbookFileName), WithoutXlsm(file), StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length == 1)
        {
            if (matches[0].UseStandardQuestionTable)
            {
                // Adding a same-edition alternative must not silently replace the
                // workbook already associated with an established report name.
                var established = ExcelMaterialNameMatcher.FindBestMatches(profiles.Where(p => !p.UseStandardQuestionTable).ToArray(), name);
                if (established.Count == 1 && CanPreserveEstablished(established[0], name) && PdfMaterialProfileStore.NormalizeName(name) !=
                    PdfMaterialProfileStore.NormalizeName(matches[0].DisplayName))
                    return (established[0], null);
            }
            return (matches[0], null);
        }

        return (null, file.Length == 0
            ? "このレポート表記名は対応表にありますが、使用Excelが未記入です。"
            : $"このレポート表記名に対応するExcelの作成設定を確認してください：{file}");
    }

    private static string WithoutXlsm(string file) => file.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
        ? file[..^5] : file;

    private static bool CanPreserveEstablished(ExcelTemplateProfile profile, string name)
    {
        var normalized = PdfMaterialProfileStore.NormalizeName(name);
        var registered = PdfMaterialProfileStore.NormalizeName(profile.DisplayName + string.Join("", profile.MaterialNameKeywords));
        if (VariantQualifiers.Any(term => normalized.Contains(term, StringComparison.Ordinal) && !registered.Contains(term, StringComparison.Ordinal)))
            return false;
        return System.Text.RegularExpressions.Regex.Matches(normalized, @"\d+訂版|第\d+版|\d+(?:st|nd|rd|th)edition")
            .All(edition => registered.Contains(edition.Value, StringComparison.Ordinal));
    }
}
