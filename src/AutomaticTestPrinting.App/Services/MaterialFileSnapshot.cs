using System.IO;

namespace AutomaticTestPrinting.App.Services;

public sealed record MaterialFileSnapshot(IReadOnlyList<string> Workbooks, IReadOnlyList<string> Pdfs)
{
    public static MaterialFileSnapshot? TryCapture(IReadOnlyCollection<string> folders)
    {
        try
        {
            var workbooks = new List<string>();
            var pdfs = new List<string>();
            var files = folders.Where(Directory.Exists)
                .SelectMany(folder => Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var path in files)
            {
                var extension = Path.GetExtension(path);
                if (extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)) workbooks.Add(path);
                if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)) pdfs.Add(path);
            }
            return new(workbooks, pdfs);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            // Keep the per-request diagnostics when a folder cannot be indexed.
            return null;
        }
    }
}
