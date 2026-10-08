namespace AutomaticTestPrinting.Core.Models;

public sealed record PdfMaterialRequest(
    string StudentName,
    string MaterialName,
    string SourcePath,
    string Pages,
    bool IsAnswer);

public sealed record PdfMaterialOutput(PdfMaterialRequest Request, string Path, int PageCount);

public sealed record PdfMaterialBatchResult(
    string OutputFolder,
    IReadOnlyList<PdfMaterialOutput> Outputs,
    string? ProblemPdfPath,
    string? AnswerPdfPath,
    IReadOnlyList<ErrorLogEntry> Errors);
