using System.Text.Json;

namespace AutomaticTestPrinting.Core.Services;

public static class ConfirmedReportMaterialNames
{
    private static readonly object Gate = new();
    private static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "AutomaticTestPrinting", "confirmed-report-material-names.json");

    private static Dictionary<string, string> Read(string path) => File.Exists(path)
        ? JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path)) ?? []
        : [];

    public static string Resolve(string original, string? path = null)
    {
        lock (Gate)
        {
            try { return Read(path ?? DefaultPath).GetValueOrDefault(original.Trim()) ?? original; }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
            {
                return original;
            }
        }
    }

    public static void Save(string original, string selected, string? path = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(original);
        ArgumentException.ThrowIfNullOrWhiteSpace(selected);
        path ??= DefaultPath;
        lock (Gate)
        {
            var names = Read(path);
            names[original.Trim()] = selected.Trim();
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, JsonSerializer.Serialize(names));
                File.Move(temporary, path, overwrite: true);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }
    }
}
