namespace AutomaticTestPrinting.Core.Models;

public sealed record RecognizedReport(
    string SourcePath,
    string StudentName,
    string NormalTestSummary,
    string StageTestSummary,
    IReadOnlyList<RecognizedReportPage> Pages,
    IReadOnlyList<NormalTestRequestCandidate>? NormalTestRequests = null)
{
    public string FileName => Path.GetFileName(SourcePath);

    public string OrientationSummary => Pages.Any(page => page.RotationDegrees == 180)
        ? "上下逆のページを自動補正しました"
        : "向きの補正はありません";

    public IReadOnlyList<NormalTestRequestCandidate> TestRequests => NormalTestRequests ?? [];
}

public sealed record NormalTestRequestCandidate(
    string MaterialName,
    string Range,
    int QuestionCount,
    int PageNumber,
    bool RequiresReview);

public sealed record RecognizedReportPage(
    int PageNumber,
    int RotationDegrees,
    string Text,
    int PixelWidth = 0,
    int PixelHeight = 0,
    IReadOnlyList<RecognizedWord>? Words = null)
{
    public IReadOnlyList<RecognizedWord> RecognizedWords => Words ?? [];
}

public sealed record RecognizedWord(
    string Text,
    double Left,
    double Top,
    double Width,
    double Height,
    int LineIndex)
{
    public double CenterX => Left + (Width / 2);

    public double CenterY => Top + (Height / 2);
}
