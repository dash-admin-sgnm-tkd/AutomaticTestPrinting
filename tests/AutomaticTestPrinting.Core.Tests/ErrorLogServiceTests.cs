using System.Text;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ErrorLogServiceTests
{
    [Fact]
    public async Task SaveAsync_WritesDetailedExcelFriendlyCsv()
    {
        var root = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-{Guid.NewGuid():N}");
        var createdAt = new DateTimeOffset(2026, 9, 26, 15, 6, 7, TimeSpan.FromHours(9));
        var entries = new[]
        {
            new ErrorLogEntry(
                createdAt,
                "PDF作成",
                "来栖凛奈",
                "report.pdf",
                2,
                "教材名",
                "101-200",
                "50",
                "System.Runtime.InteropServices.COMException",
                "0x800A03EC",
                "ExcelでPDFを作成できませんでした。",
                "1行目\n詳細,値")
        };

        try
        {
            var path = await ErrorLogService.SaveAsync(root, entries, createdAt);

            Assert.Equal(
                Path.Combine(root, "処理ログ", "エラー一覧_20260926_150607000.csv"),
                path);
            var bytes = await File.ReadAllBytesAsync(path);
            Assert.Equal([0xEF, 0xBB, 0xBF], bytes[..3]);
            var csv = Encoding.UTF8.GetString(bytes[3..]);
            Assert.Contains("PDF作成,来栖凛奈,report.pdf,2,教材名,101-200,50", csv);
            Assert.Contains("0x800A03EC", csv);
            Assert.Contains("\"1行目\n詳細,値\"", csv);
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, recursive: true);
            }
        }
    }
}
