namespace AutomaticTestPrinting.Core.Models;

public sealed record RecognizedReport(
    string SourcePath,
    string StudentName,
    string NormalTestSummary,
    string StageTestSummary,
    IReadOnlyList<RecognizedReportPage> Pages)
{
    public string FileName => Path.GetFileName(SourcePath);

    public string OrientationSummary => Pages.Any(page => page.RotationDegrees == 180)
        ? "上下逆のページを自動補正しました"
        : "向きの補正はありません";
}

public sealed record RecognizedReportPage(
    int PageNumber,
    int RotationDegrees,
    string Text);
