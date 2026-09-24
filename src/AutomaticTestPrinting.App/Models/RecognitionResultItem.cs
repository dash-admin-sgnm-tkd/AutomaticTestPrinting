using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using System.IO;

namespace AutomaticTestPrinting.App.Models;

public sealed record RecognitionResultItem(
    string SourcePath,
    string FileName,
    string StudentName,
    string NormalTestSummary,
    string NormalTestDetails,
    string ExcelPreparationDetails,
    string StageTestSummary,
    string OrientationSummary,
    bool HasError)
{
    public static RecognitionResultItem Success(RecognizedReport report, string? materialFolder) => new(
        report.SourcePath,
        report.FileName,
        report.StudentName,
        report.NormalTestSummary,
        FormatTestRequests(report.TestRequests),
        FormatExcelPreparation(report.TestRequests, materialFolder),
        report.StageTestSummary,
        report.OrientationSummary,
        false);

    public static RecognitionResultItem Failure(string sourcePath, string message) => new(
        sourcePath,
        Path.GetFileName(sourcePath),
        "読み取りに失敗しました",
        message,
        string.Empty,
        string.Empty,
        "PDFを開けるか確認してください",
        string.Empty,
        true);

    private static string FormatTestRequests(IReadOnlyList<NormalTestRequestCandidate> requests) =>
        requests.Count == 0
            ? string.Empty
            : string.Join(
                Environment.NewLine,
                requests.Select((request, index) =>
                    $"{index + 1}. {request.MaterialName} ／ 範囲 {request.Range} ／ {request.QuestionCount}問"));

    private static string FormatExcelPreparation(
        IReadOnlyList<NormalTestRequestCandidate> requests,
        string? materialFolder) =>
        requests.Count == 0
            ? string.Empty
            : string.Join(
                Environment.NewLine,
                requests.Select((request, index) =>
                    $"{index + 1}. {ExcelTemplateCatalog.Prepare(request, materialFolder).Message}"));
}
