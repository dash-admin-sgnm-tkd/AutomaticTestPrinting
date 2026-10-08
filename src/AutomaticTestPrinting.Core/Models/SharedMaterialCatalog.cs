namespace AutomaticTestPrinting.Core.Models;

public sealed record SharedMaterialEntry
{
    public string Id { get; init; } = "";
    public string Kind { get; init; } = "";
    public string Subject { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Edition { get; init; } = "";
    public IReadOnlyList<string> FileNames { get; init; } = [];
    public IReadOnlyList<string> Aliases { get; init; } = [];
    public string Notes { get; init; } = "";
    public string SourceSheet { get; init; } = "";
    public int SourceRow { get; init; }
    public string SourceSpreadsheetId { get; init; } = "";
    public string RegistrationReviewReason { get; init; } = "";
    public string SourceFileId { get; init; } = "";
    public string SourceRelativePath { get; init; } = "";
    public bool AutoEnabled { get; init; }
    public IReadOnlyList<PdfQuestionMapping> Questions { get; init; } = [];
    public string Status => Kind == "Excel"
        ? string.IsNullOrWhiteSpace(RegistrationReviewReason) ? "校舎のExcel登録設定を確認" : "情報要確認（登録は保留）"
        : AutoEnabled ? "対応表あり（実印刷未検証）" : "要確認・自動作成対象外";
    public string FileNameSummary => string.Join(" / ", FileNames);
}

public sealed record SharedMaterialDatabase
{
    public int SchemaVersion { get; init; }
    public string UpdatedAt { get; init; } = "";
    public string SourceSpreadsheetId { get; init; } = "";
    public string SourceFolderId { get; init; } = "";
    public string ScopeDescription { get; init; } = "";
    public IReadOnlyList<SharedMaterialEntry> Items { get; init; } = [];
}
