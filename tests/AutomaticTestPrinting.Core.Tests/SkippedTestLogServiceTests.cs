using System.Text;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class SkippedTestLogServiceTests
{
    [Fact]
    public async Task SaveAsync_WritesExcelFriendlyCsvWithReasons()
    {
        var root = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-{Guid.NewGuid():N}");
        var createdAt = new DateTimeOffset(2026, 9, 26, 14, 5, 6, TimeSpan.FromHours(9));
        var entries = new[]
        {
            new SkippedTestLogEntry(
                createdAt,
                "来栖凛奈",
                "report.pdf",
                2,
                "未対応,教材",
                "101-200",
                "50",
                "未対応です。\n教材登録が必要です。")
        };

        try
        {
            var path = await SkippedTestLogService.SaveAsync(root, entries, createdAt);

            Assert.Equal(
                Path.Combine(root, "処理ログ", "スキップ一覧_20260926_140506000.csv"),
                path);
            var bytes = await File.ReadAllBytesAsync(path);
            Assert.Equal([0xEF, 0xBB, 0xBF], bytes[..3]);
            var csv = Encoding.UTF8.GetString(bytes[3..]);
            Assert.Contains("来栖凛奈,report.pdf,2,\"未対応,教材\",101-200,50", csv);
            Assert.Contains("\"未対応です。\n教材登録が必要です。\"", csv);
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, recursive: true);
            }
        }
    }

    [Fact]
    public async Task SaveAsync_RejectsEmptyLog()
    {
        var exception = await Assert.ThrowsAsync<ArgumentException>(() =>
            SkippedTestLogService.SaveAsync(
                Path.GetTempPath(),
                [],
                DateTimeOffset.Now));

        Assert.Contains("スキップされた項目", exception.Message);
    }
}
