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
    bool SectionMappingUsesGridPairs = false,
    bool SourceHasSeparateDisplayNumber = false,
    double QuestionColumnWidth = 18,
    double AnswerColumnWidth = 72,
    double OutputRowHeight = 30,
    bool AutoFitOutputRows = false,
    string TeacherHeaderText = "解答",
    string StudentHeaderText = "問題",
    string? HeaderTitle = null,
    bool FitToSinglePageTall = false);

public sealed record ExcelRequestPreparation(
    bool IsSupported,
    bool IsValid,
    string Message,
    ExcelTemplateProfile? Profile = null,
    string? WorkbookPath = null,
    int? StartNumber = null,
    int? EndNumber = null);
