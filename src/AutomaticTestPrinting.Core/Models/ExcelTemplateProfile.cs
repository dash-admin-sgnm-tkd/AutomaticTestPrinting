namespace AutomaticTestPrinting.Core.Models;

public sealed record ExcelTemplateProfile
{
    public string Id { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public string WorkbookNameKeyword { get; init; } = string.Empty;
    public IReadOnlyList<string> MaterialNameKeywords { get; init; } = [];
    public int MaximumQuestionNumber { get; init; }
    public int MaximumQuestionCount { get; init; }
    public string WorkingSheetName { get; init; } = string.Empty;
    public string RangeStartCell { get; init; } = string.Empty;
    public string RangeEndCell { get; init; } = string.Empty;
    public string QuestionListSheetName { get; init; } = string.Empty;
    public string TeacherSheetName { get; init; } = string.Empty;
    public string StudentSheetName { get; init; } = string.Empty;
    public bool UseWideAnswerLayout { get; init; }
    public bool RangeUsesSectionMapping { get; init; }
    public string? SectionMappingSheetName { get; init; }
    public bool SectionMappingUsesGridPairs { get; init; }
    public bool SourceHasSeparateDisplayNumber { get; init; }
    public double QuestionColumnWidth { get; init; } = 18;
    public double AnswerColumnWidth { get; init; } = 72;
    public double OutputRowHeight { get; init; } = 30;
    public bool AutoFitOutputRows { get; init; }
    public string TeacherHeaderText { get; init; } = "解答";
    public string StudentHeaderText { get; init; } = "問題";
    public string? HeaderTitle { get; init; }
    public bool FitToSinglePageTall { get; init; }
    public bool AllowDuplicateQuestionNumbers { get; init; }
}

public sealed record ExcelRequestPreparation(
    bool IsSupported,
    bool IsValid,
    string Message,
    ExcelTemplateProfile? Profile = null,
    string? WorkbookPath = null,
    int? StartNumber = null,
    int? EndNumber = null);
