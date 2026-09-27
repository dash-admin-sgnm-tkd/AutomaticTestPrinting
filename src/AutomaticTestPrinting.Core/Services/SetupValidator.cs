using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class SetupValidator
{
    public static ValidationResult Validate(
        string? materialFolder,
        string? outputFolder,
        IReadOnlyCollection<ReportFile> reports) =>
        ValidateFolders(
            string.IsNullOrWhiteSpace(materialFolder) ? [] : [materialFolder],
            outputFolder,
            reports);

    public static ValidationResult ValidateFolders(
        IReadOnlyCollection<string> materialFolders,
        string? outputFolder,
        IReadOnlyCollection<ReportFile> reports)
    {
        if (materialFolders.Count == 0 || !materialFolders.Any(Directory.Exists))
        {
            return ValidationResult.Failure("利用できる教材フォルダーを1件以上追加してください。");
        }

        if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
        {
            return ValidationResult.Failure("出力フォルダーを選択してください。");
        }

        if (reports.Count == 0)
        {
            return ValidationResult.Failure("読み取るレポートPDFを1件以上追加してください。");
        }

        var missingReport = reports.FirstOrDefault(report => !File.Exists(report.FullPath));
        if (missingReport is not null)
        {
            return ValidationResult.Failure($"レポートが見つかりません：{missingReport.FileName}");
        }

        return ValidationResult.Success();
    }
}
