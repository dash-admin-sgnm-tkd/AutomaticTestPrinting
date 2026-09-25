using System.Reflection;
using System.Text.Json;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ExcelTemplateProfileStore
{
    public const string DefaultFileName = "materials.json";

    private const string EmbeddedResourceName =
        "AutomaticTestPrinting.Core.Configuration.materials.json";

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        AllowTrailingCommas = true,
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip
    };

    private static readonly Lazy<ExcelTemplateProfileLoadResult> DefaultProfiles =
        new(LoadDefaultCore, LazyThreadSafetyMode.ExecutionAndPublication);

    public static ExcelTemplateProfileLoadResult LoadDefault() => DefaultProfiles.Value;

    public static ExcelTemplateProfileLoadResult LoadFromFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return Failure("教材設定ファイルの場所が指定されていません。", path);
        }

        try
        {
            using var stream = File.OpenRead(path);
            return Load(stream, path);
        }
        catch (FileNotFoundException)
        {
            return Failure($"教材設定ファイルが見つかりません：{path}", path);
        }
        catch (UnauthorizedAccessException)
        {
            return Failure($"教材設定ファイルを読み取れません：{path}", path);
        }
        catch (IOException exception)
        {
            return Failure(
                $"教材設定ファイルの読み取りに失敗しました：{exception.Message}",
                path);
        }
    }

    private static ExcelTemplateProfileLoadResult LoadDefaultCore()
    {
        var externalPath = Path.Combine(AppContext.BaseDirectory, DefaultFileName);
        if (File.Exists(externalPath))
        {
            return LoadFromFile(externalPath);
        }

        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(EmbeddedResourceName);
        return stream is null
            ? Failure("内蔵の教材設定を読み込めませんでした。アプリを再配置してください。", externalPath)
            : Load(stream, "内蔵設定");
    }

    private static ExcelTemplateProfileLoadResult Load(Stream stream, string source)
    {
        try
        {
            var configuration = JsonSerializer.Deserialize<ExcelTemplateConfiguration>(
                stream,
                SerializerOptions);
            if (configuration is null)
            {
                return Failure("教材設定ファイルが空です。", source);
            }

            var validationMessage = Validate(configuration);
            return validationMessage is null
                ? new ExcelTemplateProfileLoadResult(
                    true,
                    $"教材設定を{configuration.Materials.Count}件読み込みました。",
                    configuration.Materials,
                    source)
                : Failure(validationMessage, source);
        }
        catch (JsonException exception)
        {
            return Failure(
                $"教材設定ファイルの形式が正しくありません：{exception.Message}",
                source);
        }
    }

    private static string? Validate(ExcelTemplateConfiguration configuration)
    {
        if (configuration.SchemaVersion != 1)
        {
            return $"教材設定のschemaVersionは1を指定してください（現在：{configuration.SchemaVersion}）。";
        }

        if (configuration.Materials is null || configuration.Materials.Count == 0)
        {
            return "教材設定が1件も登録されていません。";
        }

        var duplicateId = configuration.Materials
            .Where(profile => !string.IsNullOrWhiteSpace(profile.Id))
            .GroupBy(profile => profile.Id, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1)?.Key;
        if (duplicateId is not null)
        {
            return $"教材ID「{duplicateId}」が重複しています。";
        }

        foreach (var profile in configuration.Materials)
        {
            var name = string.IsNullOrWhiteSpace(profile.DisplayName)
                ? profile.Id
                : profile.DisplayName;
            if (string.IsNullOrWhiteSpace(profile.Id) ||
                string.IsNullOrWhiteSpace(profile.DisplayName) ||
                string.IsNullOrWhiteSpace(profile.WorkbookNameKeyword) ||
                profile.MaterialNameKeywords is null ||
                profile.MaterialNameKeywords.Count == 0 ||
                profile.MaterialNameKeywords.Any(string.IsNullOrWhiteSpace) ||
                profile.MaximumQuestionNumber < 1 ||
                profile.MaximumQuestionCount < 1 ||
                string.IsNullOrWhiteSpace(profile.WorkingSheetName) ||
                string.IsNullOrWhiteSpace(profile.QuestionListSheetName) ||
                string.IsNullOrWhiteSpace(profile.TeacherSheetName) ||
                string.IsNullOrWhiteSpace(profile.StudentSheetName))
            {
                return $"教材「{name}」に必須設定の不足または不正な数値があります。";
            }

            if (!profile.RangeUsesSectionMapping &&
                (string.IsNullOrWhiteSpace(profile.RangeStartCell) ||
                 string.IsNullOrWhiteSpace(profile.RangeEndCell)))
            {
                return $"教材「{name}」の範囲入力セルが設定されていません。";
            }

            if (profile.RangeUsesSectionMapping &&
                string.IsNullOrWhiteSpace(profile.SectionMappingSheetName))
            {
                return $"教材「{name}」の範囲対応表シートが設定されていません。";
            }

            if (profile.UseWideAnswerLayout &&
                (profile.QuestionColumnWidth <= 0 ||
                 profile.AnswerColumnWidth <= 0 ||
                 profile.OutputRowHeight <= 0))
            {
                return $"教材「{name}」のPDFレイアウト設定が正しくありません。";
            }
        }

        return null;
    }

    private static ExcelTemplateProfileLoadResult Failure(string message, string? source) =>
        new(false, message, [], source ?? string.Empty);
}
