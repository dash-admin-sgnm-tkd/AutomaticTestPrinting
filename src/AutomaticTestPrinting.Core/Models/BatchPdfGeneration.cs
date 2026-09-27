namespace AutomaticTestPrinting.Core.Models;

public sealed record BatchPdfGenerationResult(
    string ProblemPdfPath,
    string AnswerPdfPath,
    int TestCount);
