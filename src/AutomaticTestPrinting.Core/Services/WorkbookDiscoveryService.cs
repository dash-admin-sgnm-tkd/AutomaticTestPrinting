using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static partial class WorkbookDiscoveryService
{
    private static readonly XNamespace SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipDocumentNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    public static WorkbookDiscoveryResult Discover(
        string materialFolder,
        IReadOnlyList<ExcelTemplateProfile> registeredProfiles) =>
        Discover([materialFolder], registeredProfiles);

    public static WorkbookDiscoveryResult Discover(
        IReadOnlyCollection<string> materialFolders,
        IReadOnlyList<ExcelTemplateProfile> registeredProfiles)
    {
        ArgumentNullException.ThrowIfNull(materialFolders);
        ArgumentNullException.ThrowIfNull(registeredProfiles);

        var availableFolders = materialFolders
            .Where(folder => !string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (availableFolders.Length == 0)
        {
            throw new DirectoryNotFoundException("利用できる教材フォルダーがありません。");
        }

        var workbookPaths = availableFolders
            .SelectMany(folder => EnumerateWorkbookFiles(folder))
            .GroupBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
        var candidates = new ConcurrentBag<WorkbookRegistrationCandidate>();
        var registeredCount = 0;
        var unreadableCount = 0;
        Parallel.ForEach(
            workbookPaths,
            new ParallelOptions { MaxDegreeOfParallelism = 4 },
            workbookPath =>
        {
            var fileName = Path.GetFileName(workbookPath);
            var fileStem = Path.GetFileNameWithoutExtension(workbookPath);
            if (IsRegisteredWorkbook(fileStem, registeredProfiles))
            {
                Interlocked.Increment(ref registeredCount);
                return;
            }

            try
            {
                using var archive = ZipFile.OpenRead(workbookPath);
                var workbookMap = ReadWorkbookMap(archive);
                var matchingProfiles = registeredProfiles
                    .Where(profile => HasRequiredSheets(profile, workbookMap.Keys))
                    .OrderBy(profile => profile.RangeUsesSectionMapping)
                    .ThenBy(profile => profile.UseWideAnswerLayout)
                    .ThenByDescending(profile => RequiredSheetNames(profile).Length)
                    .ThenBy(profile => profile.Id.StartsWith(
                        "custom-",
                        StringComparison.OrdinalIgnoreCase))
                    .ThenByDescending(profile => profile.MaximumQuestionCount)
                    .ThenBy(profile => profile.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                    .ToArray();

                if (matchingProfiles.Length == 0)
                {
                    candidates.Add(new WorkbookRegistrationCandidate(
                        workbookPath,
                        fileName,
                        CreateDisplayName(fileStem),
                        [],
                        0,
                        false,
                        true,
                        "対応するシート構成を判定できませんでした"));
                    return;
                }

                var suggestedFormatId = InferFinishedLayoutId(
                    archive,
                    workbookMap,
                    fileStem,
                    matchingProfiles);
                var bestProfile = matchingProfiles.FirstOrDefault(profile => string.Equals(
                        profile.FinishedLayoutId,
                        suggestedFormatId,
                        StringComparison.OrdinalIgnoreCase))
                    ?? matchingProfiles[0];
                var maximumQuestionNumber = bestProfile.RangeUsesSectionMapping
                    ? bestProfile.MaximumQuestionNumber
                    : ReadMaximumQuestionNumber(
                        archive,
                        workbookMap,
                        bestProfile.QuestionListSheetName);
                if (maximumQuestionNumber < 1)
                {
                    maximumQuestionNumber = bestProfile.MaximumQuestionNumber;
                }

                var distinctFormats = matchingProfiles
                    .Select(CreateFormatSignature)
                    .Distinct(StringComparer.Ordinal)
                    .Count();
                var requiresReview = distinctFormats > 1;
                candidates.Add(new WorkbookRegistrationCandidate(
                    workbookPath,
                    fileName,
                    CreateDisplayName(fileStem),
                    matchingProfiles.Select(profile => profile.Id).ToArray(),
                    maximumQuestionNumber,
                    true,
                    requiresReview,
                    requiresReview && suggestedFormatId is null
                        ? "複数の完成形式に該当します。プレビューで選択してください"
                        : requiresReview
                            ? $"内容から{bestProfile.FinishedLayoutName}と判定（変更できます）"
                            : $"{bestProfile.DisplayName}と同じ形式")
                {
                    SuggestedFormatId = suggestedFormatId
                });
            }
            catch (InvalidDataException)
            {
                Interlocked.Increment(ref unreadableCount);
                candidates.Add(CreateUnreadableCandidate(workbookPath, fileName, fileStem));
            }
            catch (IOException)
            {
                Interlocked.Increment(ref unreadableCount);
                candidates.Add(CreateUnreadableCandidate(workbookPath, fileName, fileStem));
            }
            catch (UnauthorizedAccessException)
            {
                Interlocked.Increment(ref unreadableCount);
                candidates.Add(CreateUnreadableCandidate(workbookPath, fileName, fileStem));
            }
        });

        return new WorkbookDiscoveryResult(
            candidates.OrderBy(
                candidate => candidate.WorkbookPath,
                StringComparer.OrdinalIgnoreCase).ToArray(),
            registeredCount,
            unreadableCount);
    }

    private static string[] EnumerateWorkbookFiles(string materialFolder)
    {
        try
        {
            return Directory.EnumerateFiles(
                    materialFolder,
                    "*.xlsm",
                    SearchOption.AllDirectories)
                .ToArray();
        }
        catch (UnauthorizedAccessException)
        {
            return [];
        }
        catch (IOException)
        {
            return [];
        }
    }

    private static Dictionary<string, string> ReadWorkbookMap(ZipArchive archive)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("workbook.xmlがありません。");
        var relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("workbook.xml.relsがありません。");

        using var workbookStream = workbookEntry.Open();
        using var relationshipStream = relationshipsEntry.Open();
        var workbook = XDocument.Load(workbookStream);
        var relationships = XDocument.Load(relationshipStream)
            .Descendants(PackageRelationshipNamespace + "Relationship")
            .Where(element => element.Attribute("Id") is not null &&
                              element.Attribute("Target") is not null)
            .ToDictionary(
                element => (string)element.Attribute("Id")!,
                element => NormalizeWorksheetPath((string)element.Attribute("Target")!),
                StringComparer.Ordinal);

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Descendants(SpreadsheetNamespace + "sheet"))
        {
            var name = (string?)sheet.Attribute("name");
            var relationshipId = (string?)sheet.Attribute(
                RelationshipDocumentNamespace + "id");
            if (!string.IsNullOrWhiteSpace(name) &&
                !string.IsNullOrWhiteSpace(relationshipId) &&
                relationships.TryGetValue(relationshipId, out var path))
            {
                result[name] = path;
            }
        }

        return result;
    }

    private static int ReadMaximumQuestionNumber(
        ZipArchive archive,
        Dictionary<string, string> workbookMap,
        string questionListSheetName)
    {
        if (!workbookMap.TryGetValue(questionListSheetName, out var worksheetPath))
        {
            return 0;
        }

        var worksheetEntry = archive.GetEntry(worksheetPath);
        if (worksheetEntry is null)
        {
            return 0;
        }

        using var worksheetStream = worksheetEntry.Open();
        var worksheet = XDocument.Load(worksheetStream);
        var maximum = 0;
        foreach (var cell in worksheet.Descendants(SpreadsheetNamespace + "c"))
        {
            var reference = (string?)cell.Attribute("r");
            if (string.IsNullOrWhiteSpace(reference) ||
                !reference.StartsWith('B'))
            {
                continue;
            }

            var value = (string?)cell.Element(SpreadsheetNamespace + "v");
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
            {
                maximum = Math.Max(maximum, number);
            }
        }

        return maximum;
    }

    private static bool IsRegisteredWorkbook(
        string workbookFileStem,
        IReadOnlyList<ExcelTemplateProfile> profiles)
    {
        var normalizedFileName = Normalize(workbookFileStem);
        return profiles.Any(profile =>
        {
            var keyword = Normalize(profile.WorkbookNameKeyword);
            var excluded = profile.WorkbookNameExcludedKeywords.Any(item =>
                normalizedFileName.Contains(Normalize(item), StringComparison.Ordinal));
            return keyword.Length > 0 &&
                normalizedFileName.Contains(keyword, StringComparison.Ordinal) &&
                !excluded;
        });
    }

    private static bool HasRequiredSheets(
        ExcelTemplateProfile profile,
        IEnumerable<string> sheetNames)
    {
        var available = sheetNames.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return RequiredSheetNames(profile).All(available.Contains);
    }

    private static string[] RequiredSheetNames(ExcelTemplateProfile profile) =>
        new[]
        {
            profile.WorkingSheetName,
            profile.QuestionListSheetName,
            profile.TeacherSheetName,
            profile.StudentSheetName,
            profile.RangeUsesSectionMapping ? profile.SectionMappingSheetName : null
        }
        .Where(name => !string.IsNullOrWhiteSpace(name))
        .Select(name => name!)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();

    private static string CreateFormatSignature(ExcelTemplateProfile profile) => string.Join(
        '|',
        profile.FinishedLayoutId,
        profile.WorkingSheetName,
        profile.QuestionListSheetName,
        profile.TeacherSheetName,
        profile.StudentSheetName,
        profile.RangeUsesSectionMapping,
        profile.SectionMappingSheetName,
        profile.SectionMappingUsesGridPairs,
        profile.SourceHasSeparateDisplayNumber,
        profile.UseWideAnswerLayout,
        profile.AutoFitOutputRows,
        profile.FitToSinglePageTall);

    private static string? InferFinishedLayoutId(
        ZipArchive archive,
        Dictionary<string, string> workbookMap,
        string fileStem,
        IReadOnlyCollection<ExcelTemplateProfile> matchingProfiles)
    {
        var availableLayoutIds = matchingProfiles
            .Select(profile => profile.FinishedLayoutId)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        string? inferred = InferLayoutFromFileName(fileStem);
        if (inferred is null)
        {
            var questionListSheetName = matchingProfiles
                .Select(profile => profile.QuestionListSheetName)
                .FirstOrDefault(workbookMap.ContainsKey);
            inferred = questionListSheetName is null
                ? null
                : InferLayoutFromQuestionListColumns(
                    archive,
                    workbookMap[questionListSheetName]);
        }

        return inferred is not null && availableLayoutIds.Contains(inferred)
            ? inferred
            : null;
    }

    private static string? InferLayoutFromFileName(string fileStem)
    {
        var normalized = Normalize(fileStem);
        if (ContainsAny(normalized, ["語彙", "漢字", "古文単語"]))
        {
            return "japanese-vocabulary";
        }

        if (ContainsAny(normalized,
                ["英文法", "語法", "vintage", "scramble", "nextstage", "upgrade", "grammar", "大岩"]))
        {
            return "grammar-choice";
        }

        if (ContainsAny(normalized,
                ["英単語", "英熟語", "単熟語", "ターゲット", "システム英単語", "速読英単語", "leap", "鉄壁", "パス単", "stock", "sparta"]))
        {
            return "word-pair-list";
        }

        return null;
    }

    private static string? InferLayoutFromQuestionListColumns(
        ZipArchive archive,
        string worksheetPath)
    {
        var worksheetEntry = archive.GetEntry(worksheetPath);
        if (worksheetEntry is null)
        {
            return null;
        }

        using var worksheetStream = worksheetEntry.Open();
        var worksheet = XDocument.Load(worksheetStream);
        var questionWidth = ReadColumnWidth(worksheet, 3);
        var answerWidth = ReadColumnWidth(worksheet, 4);
        if (questionWidth is null || answerWidth is null)
        {
            return null;
        }

        if (questionWidth > answerWidth * 1.3)
        {
            return "grammar-choice";
        }

        return answerWidth > questionWidth * 1.3
            ? "word-pair-list"
            : null;
    }

    private static double? ReadColumnWidth(XDocument worksheet, int columnNumber)
    {
        foreach (var column in worksheet.Descendants(SpreadsheetNamespace + "col"))
        {
            if (!int.TryParse((string?)column.Attribute("min"), out var minimum) ||
                !int.TryParse((string?)column.Attribute("max"), out var maximum) ||
                columnNumber < minimum ||
                columnNumber > maximum)
            {
                continue;
            }

            if (double.TryParse(
                    (string?)column.Attribute("width"),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var width))
            {
                return width;
            }
        }

        return null;
    }

    private static bool ContainsAny(string value, IReadOnlyCollection<string> keywords) =>
        keywords.Any(keyword => value.Contains(Normalize(keyword), StringComparison.Ordinal));

    private static WorkbookRegistrationCandidate CreateUnreadableCandidate(
        string workbookPath,
        string fileName,
        string fileStem) => new(
            workbookPath,
            fileName,
            CreateDisplayName(fileStem),
            [],
            0,
            false,
            true,
            "Excelファイルの構造を読み取れませんでした");

    private static string NormalizeWorksheetPath(string target)
    {
        var normalized = target.Replace('\\', '/').TrimStart('/');
        while (normalized.StartsWith("../", StringComparison.Ordinal))
        {
            normalized = normalized[3..];
        }
        return normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
            ? normalized
            : $"xl/{normalized}";
    }

    private static string CreateDisplayName(string fileStem)
    {
        var value = fileStem.Trim().TrimStart('★', '☆');
        value = DateSuffixRegex().Replace(value, string.Empty);
        return value.Trim(' ', '_', '-');
    }

    private static string Normalize(string value)
    {
        var normalized = value.Normalize(NormalizationForm.FormKC);
        return string.Concat(normalized
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant));
    }

    [GeneratedRegex(@"[_\- ]*\d{8}(?:本部)?(?:\([^)]*\))?$", RegexOptions.IgnoreCase)]
    private static partial Regex DateSuffixRegex();
}
