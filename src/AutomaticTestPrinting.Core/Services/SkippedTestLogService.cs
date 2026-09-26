using System.Globalization;
using System.Text;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class SkippedTestLogService
{
    public static async Task<string> SaveAsync(
        string outputFolder,
        IReadOnlyCollection<SkippedTestLogEntry> entries,
        DateTimeOffset createdAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFolder);
        ArgumentNullException.ThrowIfNull(entries);

        if (entries.Count == 0)
        {
            throw new ArgumentException("スキップされた項目がありません。", nameof(entries));
        }

        var logFolder = Path.Combine(outputFolder, "処理ログ");
        Directory.CreateDirectory(logFolder);
        var path = Path.Combine(
            logFolder,
            $"スキップ一覧_{createdAt:yyyyMMdd_HHmmssfff}.csv");

        var content = BuildCsv(entries);
        await File.WriteAllTextAsync(
            path,
            content,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true),
            cancellationToken);
        return path;
    }

    internal static string BuildCsv(IEnumerable<SkippedTestLogEntry> entries)
    {
        var builder = new StringBuilder();
        builder.AppendLine("記録日時,生徒名,レポートファイル,テスト番号,教材名,範囲,問題数,スキップ理由");
        foreach (var entry in entries)
        {
            var values = new[]
            {
                entry.LoggedAt.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture),
                entry.StudentName,
                entry.ReportFileName,
                entry.TestNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                entry.MaterialName,
                entry.Range,
                entry.QuestionCount,
                entry.Reason
            };
            builder.AppendLine(string.Join(',', values.Select(Escape)));
        }

        return builder.ToString();
    }

    private static string Escape(string? value)
    {
        var text = value ?? string.Empty;
        return text.IndexOfAny([',', '"', '\r', '\n']) >= 0
            ? $"\"{text.Replace("\"", "\"\"")}\""
            : text;
    }
}
