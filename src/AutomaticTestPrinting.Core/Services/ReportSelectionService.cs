using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ReportSelectionService
{
    public static IReadOnlyList<ReportFile> CreateDistinct(
        IEnumerable<string> candidates,
        IEnumerable<ReportFile>? existing = null)
    {
        var knownPaths = new HashSet<string>(
            existing?.Select(report => report.FullPath) ?? [],
            StringComparer.OrdinalIgnoreCase);

        var reports = new List<ReportFile>();

        foreach (var candidate in candidates)
        {
            if (!string.Equals(Path.GetExtension(candidate), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var fullPath = Path.GetFullPath(candidate);
            if (knownPaths.Add(fullPath))
            {
                reports.Add(new ReportFile(fullPath));
            }
        }

        return reports;
    }
}
