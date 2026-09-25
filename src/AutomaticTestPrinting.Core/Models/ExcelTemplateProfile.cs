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
    string QuestionListSheetName,
    string TeacherSheetName,
    string StudentSheetName,
    bool UseWideAnswerLayout = false,
    bool RangeUsesSectionMapping = false,
    string? SectionMappingSheetName = null,
    double QuestionColumnWidth = 18,
    double AnswerColumnWidth = 72,
    double OutputRowHeight = 30);

public sealed record ExcelRequestPreparation(
    bool IsSupported,
    bool IsValid,
    string Message,
    ExcelTemplateProfile? Profile = null,
    string? WorkbookPath = null,
    int? StartNumber = null,
    int? EndNumber = null);
