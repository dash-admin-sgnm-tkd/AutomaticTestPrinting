namespace AutomaticTestPrinting.Core.Models;

public sealed record ExcelMaterialFormatOption(
    string Id,
    string DisplayName,
    ExcelTemplateProfile TemplateProfile,
    IReadOnlyList<string> SourceProfileIds,
    string PreviewDescription,
    string PreviewQuestion,
    string PreviewAnswer);
