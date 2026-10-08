namespace AutomaticTestPrinting.Core.Models;

public sealed record PdfQuestionMapping(
    int Number, string FileName, string Pages, string AnswerFileName = "", string AnswerPages = "", int? ExpectedPageCount = null);

public sealed record PdfMaterialProfile(
    string Id, string DisplayName, IReadOnlyList<string> Aliases, IReadOnlyList<PdfQuestionMapping> Questions, bool IsStageTest = false);

public sealed record ResolvedPdfQuestion(int Number, string Path, string Pages, string? AnswerPath, string AnswerPages, int? ExpectedPageCount = null);

public sealed record PdfRequestPreparation(
    bool IsSupported, bool IsValid, string Message, IReadOnlyList<ResolvedPdfQuestion>? Questions = null);

public sealed record TestGenerationPlan(
    string StudentName, string MaterialName, string Range, int QuestionCount,
    string SourceReportFileName, int TestNumber,
    ExcelPdfGenerationRequest? ExcelRequest = null, IReadOnlyList<ResolvedPdfQuestion>? PdfQuestions = null,
    string? SourceReportPath = null);
