using AutomaticTestPrinting.Core.Models;
using System.IO;

namespace AutomaticTestPrinting.App.Models;

public sealed record RecognitionResultItem(
    string SourcePath,
    string FileName,
    string StudentName,
    string NormalTestSummary,
    string StageTestSummary,
    string OrientationSummary,
    bool HasError)
{
    public static RecognitionResultItem Success(RecognizedReport report) => new(
        report.SourcePath,
        report.FileName,
        report.StudentName,
        report.NormalTestSummary,
        report.StageTestSummary,
        report.OrientationSummary,
        false);

    public static RecognitionResultItem Failure(string sourcePath, string message) => new(
        sourcePath,
        Path.GetFileName(sourcePath),
        "読み取りに失敗しました",
        message,
        "PDFを開けるか確認してください",
        string.Empty,
        true);
}
