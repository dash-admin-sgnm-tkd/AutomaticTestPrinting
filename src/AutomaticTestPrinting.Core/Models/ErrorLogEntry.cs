namespace AutomaticTestPrinting.Core.Models;

public sealed record ErrorLogEntry(
    DateTimeOffset LoggedAt,
    string Stage,
    string StudentName,
    string ReportFileName,
    int? TestNumber,
    string MaterialName,
    string Range,
    string QuestionCount,
    string ErrorType,
    string ErrorCode,
    string UserMessage,
    string TechnicalDetails);
