namespace AutomaticTestPrinting.Core.Models;

public sealed record ExcelPdfGenerationRequest(
    string StudentName,
    string MaterialName,
    int QuestionCount,
    int StartNumber,
    int EndNumber,
    string WorkbookPath,
    string OutputFolder,
    ExcelTemplateProfile Profile,
    string SourceReportFileName,
    int TestNumber);

public sealed record ExcelPdfGenerationResult(
    string ProblemPdfPath,
    string AnswerPdfPath);
