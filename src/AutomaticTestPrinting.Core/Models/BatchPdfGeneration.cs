namespace AutomaticTestPrinting.Core.Models;

public sealed record GeneratedReportTest(string SourceReportPath, ExcelPdfGenerationResult Pdf);

public sealed record ReportBatchPdfGenerationResult(string PrintPdfPath, int TestCount, int ReportCount);

public sealed record BatchPdfGenerationResult(
    string ProblemPdfPath,
    string? AnswerPdfPath,
    int TestCount);
