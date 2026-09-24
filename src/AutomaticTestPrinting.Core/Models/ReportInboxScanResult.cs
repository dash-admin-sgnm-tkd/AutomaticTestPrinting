namespace AutomaticTestPrinting.Core.Models;

public sealed record ReportInboxScanResult(
    IReadOnlyList<ReportFile> Reports,
    int WaitingForSyncCount);
