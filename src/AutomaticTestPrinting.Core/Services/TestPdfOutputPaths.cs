using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class TestPdfOutputPaths
{
    public static ExcelPdfGenerationResult FindAvailable(
        string outputFolder, string studentName, string materialName, string range, int questionCount,
        bool includeAnswers)
    {
        var baseName = $"{Sanitize(studentName, "氏名未確認")}_{Sanitize(materialName, "教材")}_{Sanitize(range, "範囲未確認")}_{questionCount}問";
        if (baseName.Length > 120)
        {
            baseName = baseName[..120];
        }

        for (var suffix = 1; ; suffix++)
        {
            var suffixText = suffix == 1 ? string.Empty : $"_{suffix}";
            var problem = Path.Combine(outputFolder, $"{baseName}{suffixText}_問題.pdf");
            var answer = Path.Combine(outputFolder, $"{baseName}{suffixText}_解答.pdf");
            if (!File.Exists(problem) && !File.Exists(answer))
            {
                return new(problem, includeAnswers ? answer : null);
            }
        }
    }

    private static string Sanitize(string value, string fallback)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitized = new string(value.Trim()
            .Select(character => invalidCharacters.Contains(character) ? '_' : character).ToArray());
        return string.IsNullOrWhiteSpace(sanitized) ? fallback : sanitized;
    }
}
