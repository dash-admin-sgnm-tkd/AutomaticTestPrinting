namespace AutomaticTestPrinting.Core.Models;

public sealed record SkippedTestLogEntry(
    DateTimeOffset LoggedAt,
    string StudentName,
    string ReportFileName,
    int? TestNumber,
    string MaterialName,
    string Range,
    string QuestionCount,
    string Reason);
