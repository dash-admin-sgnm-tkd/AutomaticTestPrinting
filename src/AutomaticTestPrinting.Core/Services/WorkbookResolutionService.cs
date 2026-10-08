using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public sealed record WorkbookChoice(string RelativePath, string Sha256);
public sealed record WorkbookResolution(string? Path, IReadOnlyList<string> Candidates, string Message);

public static class WorkbookResolutionService
{
    private static readonly ConcurrentDictionary<string, (long Length, DateTime Modified, string Hash)> HashCache =
        new(StringComparer.OrdinalIgnoreCase);

    public static string GetHash(string path)
    {
        var before = new FileInfo(path);
        var length = before.Length;
        var modified = before.LastWriteTimeUtc;
        if (HashCache.TryGetValue(path, out var cached) && cached.Length == length && cached.Modified == modified)
            return cached.Hash;
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var after = new FileInfo(path);
        if (after.Length != length || after.LastWriteTimeUtc != modified)
            throw new IOException("Excelの内容が照合中に変わりました。もう一度確認してください。");
        HashCache[path] = (length, modified, hash);
        return hash;
    }

    public static WorkbookResolution Resolve(IReadOnlyCollection<string> roots, ExcelTemplateProfile profile,
        WorkbookChoice? choice = null, IReadOnlyList<string>? knownCandidates = null)
    {
        var matches = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var root in roots.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!Directory.Exists(root)) continue;
                var paths = knownCandidates is null
                    ? Directory.EnumerateFiles(root, "*.xlsm", SearchOption.AllDirectories)
                    : knownCandidates.Where(path => GetRelativePath(root, path) is not null && File.Exists(path));
                foreach (var path in paths)
                {
                    var exact = !string.IsNullOrWhiteSpace(profile.WorkbookFileName)
                        ? string.Equals(System.IO.Path.GetFileName(path), profile.WorkbookFileName, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(System.IO.Path.GetFileNameWithoutExtension(path), profile.WorkbookNameKeyword, StringComparison.OrdinalIgnoreCase);
                    if (exact && !profile.WorkbookNameExcludedKeywords.Any(keyword =>
                            Normalize(System.IO.Path.GetFileNameWithoutExtension(path)).Contains(Normalize(keyword), StringComparison.Ordinal)))
                        matches.Add(System.IO.Path.GetFullPath(path));
                }
            }
            var candidates = MaterialLocationPolicy.SelectCandidates(matches);
            if (choice is not null)
            {
                var selected = candidates.Where(path => roots.Any(root =>
                    string.Equals(GetRelativePath(root, path), choice.RelativePath, StringComparison.OrdinalIgnoreCase))).ToArray();
                if (selected.Length > 0)
                {
                    var unchanged = selected.Where(path => GetHash(path) == choice.Sha256).ToArray();
                    if (unchanged.Length > 0)
                        return new(unchanged[0], candidates, "保存した使用Excelを選択しました。");
                }
                return new(null, candidates, "保存したExcelが移動・更新されたか、参照対象外（国立校など）になっています。「使用Excelを選ぶ」で再確認してください。別ファイルへ自動変更しません。");
            }
            if (candidates.Length <= 1) return new(candidates.SingleOrDefault(), candidates,
                candidates.Length == 0 && matches.Count > 0 ? "同名Excelは国立校フォルダー内にしかありません。国立校側は参照しません。" : "");
            if (candidates.Select(GetHash).Distinct(StringComparer.Ordinal).Take(2).Count() == 1)
                return new(candidates[0], candidates, $"同名の{candidates.Length}件は内容が同一のため、コピーを自動選択しました。");
            return new(null, candidates, $"同名で内容の異なるExcelが{candidates.Length}件あります。「使用Excelを選ぶ」で一度選択してください。");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return new(null, MaterialLocationPolicy.SelectCandidates(matches),
                $"Excelの照合に失敗しました。自動選択は停止します：{exception.Message}");
        }
    }

    public static string? GetRelativePath(string root, string path)
    {
        var relative = System.IO.Path.GetRelativePath(System.IO.Path.GetFullPath(root), System.IO.Path.GetFullPath(path));
        return System.IO.Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + System.IO.Path.DirectorySeparatorChar, StringComparison.Ordinal)
            ? null : relative;
    }

    private static string Normalize(string value) => string.Concat(value.Normalize(System.Text.NormalizationForm.FormKC)
        .Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant));
}

public static class WorkbookChoiceStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new() { WriteIndented = true };
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "AutomaticTestPrinting", "workbook-choices.json");

    public static Dictionary<string, WorkbookChoice> Load(string? path = null)
    {
        path ??= DefaultPath;
        if (!File.Exists(path)) return new(StringComparer.OrdinalIgnoreCase);
        var choices = JsonSerializer.Deserialize<Dictionary<string, WorkbookChoice>>(File.ReadAllText(path))
            ?? throw new InvalidDataException("使用Excelの保存設定が空です。");
        if (choices.Any(pair => string.IsNullOrWhiteSpace(pair.Key) || pair.Value is null ||
                string.IsNullOrWhiteSpace(pair.Value.RelativePath) || Path.IsPathRooted(pair.Value.RelativePath) ||
                pair.Value.RelativePath.Split(['/', '\\']).Any(part => part == "..") ||
                pair.Value.Sha256 is null || pair.Value.Sha256.Length != 64 || !pair.Value.Sha256.All(Uri.IsHexDigit)))
            throw new InvalidDataException("使用Excelの保存設定が正しくありません。相対パスと内容照合値を確認してください。");
        return new(choices, StringComparer.OrdinalIgnoreCase);
    }

    public static void Save(string profileId, WorkbookChoice choice, string? path = null)
    {
        path ??= DefaultPath;
        var choices = Load(path);
        choices[profileId] = choice;
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temporary, JsonSerializer.Serialize(choices, SerializerOptions));
            File.Move(temporary, path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
