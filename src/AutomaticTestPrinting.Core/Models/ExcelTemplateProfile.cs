namespace AutomaticTestPrinting.Core.Models;

public sealed record ExcelTemplateProfile(
    string Id,
    string DisplayName,
    string WorkbookNameKeyword,
    IReadOnlyList<string> MaterialNameKeywords,
    int MaximumQuestionNumber,
    int MaximumQuestionCount,
    string WorkingSheetName,
    string RangeStartCell,
    string RangeEndCell,
    string TeacherSheetName,
    string StudentSheetName);

public sealed record ExcelRequestPreparation(
    bool IsSupported,
    bool IsValid,
    string Message,
    ExcelTemplateProfile? Profile = null,
    string? WorkbookPath = null,
    int? StartNumber = null,
    int? EndNumber = null);
