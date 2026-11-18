using System.Text;
using System.Text.Json;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class PdfMaterialProfileStore
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true, PropertyNameCaseInsensitive = true };
    public static string UserConfigurationPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AutomaticTestPrinting", "pdf-materials.json");

    public static IReadOnlyList<PdfMaterialProfile> Load(string? path = null)
    {
        var useSharedCatalog = path is null;
        path ??= UserConfigurationPath;
        var profiles = File.Exists(path) ? JsonSerializer.Deserialize<List<PdfMaterialProfile>>(File.ReadAllText(path), Options)
            ?? throw new InvalidDataException("PDF教材設定が空です。") : [];
        Validate(profiles);
        return useSharedCatalog ? MergeProfiles(SharedMaterialCatalog.CreatePdfProfiles(), profiles) : profiles;
    }

    public static IReadOnlyList<PdfMaterialProfile> MergeProfiles(IReadOnlyList<PdfMaterialProfile> shared,
        IReadOnlyList<PdfMaterialProfile> user)
    {
        Validate(shared);
        Validate(user);
        var userNames = user.SelectMany(p => p.Aliases.Append(p.DisplayName)).Select(NormalizeName).ToHashSet();
        var merged = shared.Where(p => !user.Any(u => u.Id == p.Id) &&
            !p.Aliases.Append(p.DisplayName).Any(a => userNames.Contains(NormalizeName(a)))).Concat(user).ToArray();
        Validate(merged);
        return merged;
    }

    public static void Save(IReadOnlyList<PdfMaterialProfile> profiles, string? path = null)
    {
        Validate(profiles);
        if (path is null)
        {
            // Do not freeze unchanged bundled entries into the user override file.
            var shared = SharedMaterialCatalog.CreatePdfProfiles().ToDictionary(p => p.Id);
            profiles = profiles.Where(p => !shared.TryGetValue(p.Id, out var original) ||
                JsonSerializer.Serialize(p, Options) != JsonSerializer.Serialize(original, Options)).ToArray();
        }
        path ??= UserConfigurationPath;
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var temporary = path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(profiles, Options));
            if (File.Exists(path)) File.Copy(path, path + ".backup", overwrite: true);
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    public static string NormalizeName(string value)
    {
        var name = string.Concat(value.Normalize(NormalizationForm.FormKC)
            .Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant));
        // Report subject labels include 国語（現代文/古文/漢文）; catalog keys use the specific subject.
        // Only normalize that subject prefix, keeping year, level, and attempt intact.
        foreach (var subject in new[] { "現代文", "古文", "漢文" })
        {
            var prefix = "段階突破国語" + subject;
            if (name.StartsWith(prefix, StringComparison.Ordinal))
                return "段階突破" + subject + name[prefix.Length..];
        }
        return name;
    }

    public static void Validate(IReadOnlyList<PdfMaterialProfile> profiles)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var aliases = new HashSet<string>(StringComparer.Ordinal);
        foreach (var profile in profiles)
        {
            if (profile is null || string.IsNullOrWhiteSpace(profile.Id) || !ids.Add(profile.Id) ||
                string.IsNullOrWhiteSpace(profile.DisplayName) || profile.Aliases is null || profile.Questions is null || profile.Questions.Count == 0 ||
                profile.Aliases.Any(string.IsNullOrWhiteSpace) || profile.Questions.Any(question => question is null))
                throw new InvalidDataException("PDF教材のID・教材名・対応表を確認してください。");
            foreach (var alias in profile.Aliases.Append(profile.DisplayName).DistinctBy(NormalizeName))
            {
                if (string.IsNullOrWhiteSpace(alias) || NormalizeName(alias).Length == 0 || !aliases.Add(NormalizeName(alias)))
                    throw new InvalidDataException($"PDF教材名／別名が他の教材と重複しています：{alias}");
            }
            var numbers = new HashSet<int>();
            if (profile.IsStageTest && profile.AllowSparseRanges)
                throw new InvalidDataException("段階突破には練習問題の欠番指定を使用できません。");
            if (profile.IsStageTest && (profile.Questions.Count != 1 || profile.Questions[0].Number != 1 ||
                !string.IsNullOrWhiteSpace(profile.Questions[0].Pages) || !string.IsNullOrWhiteSpace(profile.Questions[0].AnswerFileName)))
                throw new InvalidDataException("段階突破は問題番号1・ページ空欄（全ページ）・解答欄空欄で登録してください。");
            var withAnswers = profile.Questions.Count(question => !string.IsNullOrWhiteSpace(question.AnswerFileName));
            if (withAnswers != 0 && withAnswers != profile.Questions.Count)
                throw new InvalidDataException("解答は全問に登録するか、全問空欄にしてください。");
            foreach (var question in profile.Questions)
            {
                if (question.ExpectedPageCount is <= 0)
                    throw new InvalidDataException("照合するPDF総ページ数は1以上を指定してください。");
                if (question.Number < 1 || !numbers.Add(question.Number) || question.Pages is null || question.AnswerFileName is null || question.AnswerPages is null)
                    throw new InvalidDataException("問題番号は1以上、重複なしで登録してください。");
                ValidateFileName(question.FileName);
                if (!string.IsNullOrWhiteSpace(question.AnswerFileName)) ValidateFileName(question.AnswerFileName);
            }
        }
    }

    private static void ValidateFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name != Path.GetFileName(name) ||
            name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || !name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException($"PDFのファイル名のみを指定してください（パスは不要）：{name}");
    }
}
