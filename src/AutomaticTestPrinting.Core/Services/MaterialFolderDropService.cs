namespace AutomaticTestPrinting.Core.Services;

public sealed record MaterialFolderDropResult(IReadOnlyList<string> AddedFolders, int DuplicateCount, int IgnoredCount);

public static class MaterialFolderDropService
{
    public static MaterialFolderDropResult Plan(IEnumerable<string> paths, IReadOnlyCollection<string> existingFolders)
    {
        var known = existingFolders.Select(Normalize).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = new List<string>();
        var duplicates = 0;
        var ignored = 0;
        foreach (var path in paths)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) { ignored++; continue; }
                var normalized = Normalize(path);
                if (!known.Add(normalized)) { duplicates++; continue; }
                added.Add(normalized);
            }
            catch (Exception exception) when (exception is ArgumentException or IOException or UnauthorizedAccessException or NotSupportedException)
            {
                ignored++;
            }
        }
        return new(added, duplicates, ignored);
    }

    private static string Normalize(string path) => Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
}
