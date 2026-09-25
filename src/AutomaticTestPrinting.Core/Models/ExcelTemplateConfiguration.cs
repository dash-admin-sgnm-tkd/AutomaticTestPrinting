namespace AutomaticTestPrinting.Core.Models;

public sealed record ExcelTemplateConfiguration
{
    public int SchemaVersion { get; init; } = 1;
    public IReadOnlyList<ExcelTemplateProfile> Materials { get; init; } = [];
}

public sealed record ExcelTemplateProfileLoadResult(
    bool IsValid,
    string Message,
    IReadOnlyList<ExcelTemplateProfile> Profiles,
    string Source);
