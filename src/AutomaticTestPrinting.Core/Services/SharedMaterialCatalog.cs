using System.Reflection;
using System.Text.Json;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class SharedMaterialCatalog
{
    public const string FileName = "shared-material-catalog.json";
    private static readonly JsonSerializerOptions Options = new() { PropertyNameCaseInsensitive = true };
    private static readonly Lazy<SharedMaterialDatabase> Database = new(LoadCore);
    public static SharedMaterialDatabase Load() => Database.Value;

    public static SharedMaterialDatabase Parse(string json)
    {
        var database = JsonSerializer.Deserialize<SharedMaterialDatabase>(json, Options)
            ?? throw new InvalidDataException("共通教材カタログが空です。");
        if (database.SchemaVersion != 1 || database.Items is null || database.Items.Count == 0)
            throw new InvalidDataException("共通教材カタログの版・教材一覧を確認してください。");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in database.Items)
        {
            if (item is null || string.IsNullOrWhiteSpace(item.Id) || !ids.Add(item.Id) ||
                string.IsNullOrWhiteSpace(item.DisplayName) || item.Kind is not ("Excel" or "PDF" or "段階突破") ||
                item.FileNames is null || item.Aliases is null || item.Questions is null || item.Questions.Any(q => q is null) ||
                item.Aliases.Any(string.IsNullOrWhiteSpace) || item.FileNames.Any(name =>
                    string.IsNullOrWhiteSpace(name) || name != Path.GetFileName(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0))
                throw new InvalidDataException("共通教材カタログのID・種類・教材名・ファイル名を確認してください。");
            if (item.PdfVariants is null || (item.PdfVariants.Count > 0 &&
                (item.Kind != "PDF" || item.Questions.Count > 0 || item.PdfVariants.Any(p => p is null || p.IsStageTest))))
                throw new InvalidDataException("用途別PDFの登録内容を確認してください。");
            if (item.PdfVariants.Count > 0) PdfMaterialProfileStore.Validate(item.PdfVariants);
            if (item.AutoEnabled && (item.Kind == "Excel" || (item.Questions.Count == 0 && item.PdfVariants.Count == 0)))
                throw new InvalidDataException($"未確認の教材を自動有効化できません：{item.DisplayName}");
            if (item.AutoEnabled && item.Kind == "段階突破" &&
                (item.Questions.Count != 1 || item.Questions[0].Number != 1 || item.Questions[0].Pages != "" ||
                 item.Questions[0].AnswerFileName != ""))
                throw new InvalidDataException("段階突破は全ページを一つのテストとして扱ってください。");
            var questions = item.Questions.Concat(item.PdfVariants.SelectMany(p=>p.Questions));
            if (item.AutoEnabled && questions.Any(q => !item.FileNames.Contains(q.FileName, StringComparer.OrdinalIgnoreCase)))
                throw new InvalidDataException("対応表のPDFが教材のファイル一覧にありません。");
            if (item.AutoEnabled && item.Kind == "PDF")
                foreach (var question in questions)
                {
                    if (question.ExpectedPageCount is null)
                        throw new InvalidDataException("共通PDF教材には照合用の総ページ数が必要です。");
                    PdfMaterialGenerationService.ParsePages(question.Pages, question.ExpectedPageCount.Value);
                }
        }
        PdfMaterialProfileStore.Validate(CreatePdfProfiles(database));
        return database;
    }

    public static IReadOnlyList<PdfMaterialProfile> CreatePdfProfiles(SharedMaterialDatabase? database = null) =>
        (database ?? Load()).Items.Where(item => item.AutoEnabled && item.Kind != "Excel")
            .SelectMany(item => item.PdfVariants.Count > 0 ? item.PdfVariants :
                (IReadOnlyList<PdfMaterialProfile>)[new(item.Id, item.DisplayName, item.Aliases, item.Questions,
                    IsStageTest: item.Kind == "段階突破", AllowSparseRanges: item.AllowSparseRanges)]).ToArray();

    private static SharedMaterialDatabase LoadCore()
    {
        var primary = LoadPrimary();
        var excelJson = ReadAdditionalCatalog("additional-excel-catalog.json");
        if (excelJson is not null) primary = MergeExcelCatalog(primary, Parse(excelJson));
        const string additionalName = "additional-stage-catalog.json";
        var additionalPath = Path.Combine(AppContext.BaseDirectory, additionalName);
        string? json = File.Exists(additionalPath) ? File.ReadAllText(additionalPath) : null;
        if (json is null)
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(
                "AutomaticTestPrinting.Core.Configuration." + additionalName);
            if (stream is null) return primary;
            using var reader = new StreamReader(stream);
            json = reader.ReadToEnd();
        }
        var additional = Parse(json);
        if (additional.Items.Any(item => item.Kind != "段階突破"))
            throw new InvalidDataException("追加段階突破カタログに別種の教材が含まれています。");
        var existingNames = primary.Items.SelectMany(item => item.Aliases.Append(item.DisplayName))
            .Select(PdfMaterialProfileStore.NormalizeName).ToHashSet(StringComparer.Ordinal);
        var newEntries = additional.Items.Where(item => !item.Aliases.Append(item.DisplayName)
            .Any(name => existingNames.Contains(PdfMaterialProfileStore.NormalizeName(name)))).ToArray();
        return Parse(JsonSerializer.Serialize(primary with
        {
            UpdatedAt = string.CompareOrdinal(primary.UpdatedAt, additional.UpdatedAt) >= 0
                ? primary.UpdatedAt : additional.UpdatedAt,
            ScopeDescription = primary.ScopeDescription + " ＋ 追加段階突破フォルダー",
            Items = [.. primary.Items, .. newEntries]
        }));
    }

    public static SharedMaterialDatabase MergeExcelCatalog(SharedMaterialDatabase primary, SharedMaterialDatabase additional)
    {
        if (additional.Items.Any(item => item.Kind != "Excel" || item.AutoEnabled))
            throw new InvalidDataException("追加Excelカタログは未検証のExcel情報のみを扱います。");
        var items = primary.Items.ToList();
        foreach (var incoming in additional.Items)
        {
            var matches = items.Select((entry, index) => (entry, index)).Where(pair =>
                pair.entry.Kind == "Excel" && string.IsNullOrWhiteSpace(incoming.RegistrationReviewReason) &&
                incoming.FileNames.Any(file => pair.entry.FileNames.Contains(file, StringComparer.OrdinalIgnoreCase)))
                .ToArray();
            if (matches.Length == 1)
            {
                // Keep the original scope, identity and settings; attach only source observations.
                var (existing, index) = matches[0];
                var observation = incoming.Notes;
                items[index] = existing with
                {
                    Notes = existing.Notes.Contains(observation, StringComparison.Ordinal)
                        ? existing.Notes : existing.Notes + "\n\n追加調査情報：\n" + observation
                };
            }
            else if (!items.Any(item => item.Id == incoming.Id)) items.Add(incoming);
        }
        return Parse(JsonSerializer.Serialize(primary with
        {
            UpdatedAt = string.CompareOrdinal(primary.UpdatedAt, additional.UpdatedAt) >= 0
                ? primary.UpdatedAt : additional.UpdatedAt,
            ScopeDescription = primary.ScopeDescription.Contains("Excelシート", StringComparison.Ordinal)
                ? primary.ScopeDescription : primary.ScopeDescription + " ＋ Excelシートの追加調査情報",
            Items = items
        }));
    }

    private static string? ReadAdditionalCatalog(string name)
    {
        var external = Path.Combine(AppContext.BaseDirectory, name);
        if (File.Exists(external)) return File.ReadAllText(external);
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(
            "AutomaticTestPrinting.Core.Configuration." + name);
        if (stream is null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static SharedMaterialDatabase LoadPrimary()
    {
        var external = Path.Combine(AppContext.BaseDirectory, FileName);
        if (File.Exists(external)) return Parse(File.ReadAllText(external));
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(
            "AutomaticTestPrinting.Core.Configuration.shared-material-catalog.json")
            ?? throw new InvalidDataException("内蔵の共通教材カタログを読み込めません。");
        using var reader = new StreamReader(stream);
        return Parse(reader.ReadToEnd());
    }
}
