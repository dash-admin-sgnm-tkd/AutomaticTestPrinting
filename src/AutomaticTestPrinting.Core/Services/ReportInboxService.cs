using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ReportInboxService
{
    private static readonly TimeSpan MinimumFileAge = TimeSpan.FromSeconds(3);

    public static ReportInboxScanResult Scan(
        string? inboxFolder,
        IEnumerable<ReportFile>? existingReports = null,
        DateTime? utcNow = null)
    {
        if (string.IsNullOrWhiteSpace(inboxFolder) || !Directory.Exists(inboxFolder))
        {
            return new ReportInboxScanResult([], 0);
        }

        var now = utcNow ?? DateTime.UtcNow;
        var readyPaths = new List<string>();
        var waitingForSyncCount = 0;

        try
        {
            foreach (var path in Directory.EnumerateFiles(
                         inboxFolder,
                         "*.pdf",
                         SearchOption.TopDirectoryOnly))
            {
                if (IsReady(path, now))
                {
                    readyPaths.Add(path);
                }
                else
                {
                    waitingForSyncCount++;
                }
            }
        }
        catch (UnauthorizedAccessException)
        {
            return new ReportInboxScanResult([], 0);
        }
        catch (IOException)
        {
            return new ReportInboxScanResult([], 0);
        }

        var reports = ReportSelectionService.CreateDistinct(readyPaths, existingReports);
        return new ReportInboxScanResult(reports, waitingForSyncCount);
    }

    private static bool IsReady(string path, DateTime utcNow)
    {
        try
        {
            var file = new FileInfo(path);
            if (file.Length == 0 || utcNow - file.LastWriteTimeUtc < MinimumFileAge)
            {
                return false;
            }

            using var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read);
            return stream.Length > 0;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }
}
