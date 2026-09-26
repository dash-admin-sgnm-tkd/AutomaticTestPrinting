using System.Globalization;
using System.Text;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ErrorLogService
{
    public static async Task<string> SaveAsync(
        string outputFolder,
        IReadOnlyCollection<ErrorLogEntry> entries,
        DateTimeOffset createdAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFolder);
        ArgumentNullException.ThrowIfNull(entries);

        if (entries.Count == 0)
        {
            throw new ArgumentException("記録するエラーがありません。", nameof(entries));
        }

        var logFolder = Path.Combine(outputFolder, "処理ログ");
        Directory.CreateDirectory(logFolder);
        var path = Path.Combine(logFolder, $"エラー一覧_{createdAt:yyyyMMdd_HHmmssfff}.csv");
        await File.WriteAllTextAsync(
            path,
            BuildCsv(entries),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true),
            cancellationToken);
        return path;
    }

    internal static string BuildCsv(IEnumerable<ErrorLogEntry> entries)
    {
        var builder = new StringBuilder();
        builder.AppendLine(
            "記録日時,発生工程,生徒名,レポートファイル,テスト番号,教材名,範囲,問題数,エラー種別,エラーコード,画面表示内容,技術的詳細");
        foreach (var entry in entries)
        {
            var values = new[]
            {
                entry.LoggedAt.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture),
                entry.Stage,
                entry.StudentName,
                entry.ReportFileName,
                entry.TestNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                entry.MaterialName,
                entry.Range,
                entry.QuestionCount,
                entry.ErrorType,
                entry.ErrorCode,
                entry.UserMessage,
                entry.TechnicalDetails
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
