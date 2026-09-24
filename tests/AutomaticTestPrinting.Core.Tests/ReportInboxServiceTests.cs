using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ReportInboxServiceTests : IDisposable
{
    private readonly string _directory = Path.Combine(
        Path.GetTempPath(),
        $"AutomaticTestPrinting-Inbox-{Guid.NewGuid():N}");

    [Fact]
    public void Scan_ReturnsOnlyReadyNewPdfFiles()
    {
        Directory.CreateDirectory(_directory);
        var readyPath = CreateFile("ready.pdf", "%PDF-ready", DateTime.UtcNow.AddSeconds(-10));
        var existingPath = CreateFile("existing.pdf", "%PDF-existing", DateTime.UtcNow.AddSeconds(-10));
        CreateFile("syncing.pdf", "%PDF-syncing", DateTime.UtcNow);
        CreateFile("empty.pdf", string.Empty, DateTime.UtcNow.AddSeconds(-10));
        CreateFile("note.txt", "not a pdf", DateTime.UtcNow.AddSeconds(-10));

        var result = ReportInboxService.Scan(
            _directory,
            [new ReportFile(existingPath)],
            DateTime.UtcNow);

        var report = Assert.Single(result.Reports);
        Assert.Equal(readyPath, report.FullPath);
        Assert.Equal(2, result.WaitingForSyncCount);
    }

    [Fact]
    public void Scan_ReturnsEmptyWhenFolderDoesNotExist()
    {
        var result = ReportInboxService.Scan(_directory);

        Assert.Empty(result.Reports);
        Assert.Equal(0, result.WaitingForSyncCount);
    }

    private string CreateFile(string name, string content, DateTime lastWriteTimeUtc)
    {
        var path = Path.Combine(_directory, name);
        File.WriteAllText(path, content);
        File.SetLastWriteTimeUtc(path, lastWriteTimeUtc);
        return path;
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory))
        {
            Directory.Delete(_directory, true);
        }

        GC.SuppressFinalize(this);
    }
}
