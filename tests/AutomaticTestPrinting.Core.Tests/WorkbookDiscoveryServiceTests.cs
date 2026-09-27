using System.IO.Compression;
using System.Globalization;
using System.Text;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class WorkbookDiscoveryServiceTests : IDisposable
{
    private readonly string _directory = Path.Combine(
        Path.GetTempPath(),
        $"AutomaticTestPrinting-Discovery-{Guid.NewGuid():N}");

    [Fact]
    public void Discover_FindsUnregisteredWorkbookAndReadsMaximumNumber()
    {
        Directory.CreateDirectory(_directory);
        CreateWorkbook(
            Path.Combine(_directory, "★新しい英単語_20260927本部.xlsm"),
            ["作業シート", "問題解答リスト", "講師用", "生徒用"],
            [1, 80, 123]);

        var result = WorkbookDiscoveryService.Discover(_directory, [CreateProfile()]);

        var candidate = Assert.Single(result.Candidates);
        Assert.True(candidate.CanRegister);
        Assert.Equal("新しい英単語", candidate.SuggestedDisplayName);
        Assert.Equal(123, candidate.SuggestedMaximumQuestionNumber);
        Assert.Contains("standard", candidate.MatchingProfileIds);
        Assert.Equal(0, result.RegisteredWorkbookCount);
    }

    [Fact]
    public void Discover_SkipsWorkbookAlreadyCoveredByProfileKeyword()
    {
        Directory.CreateDirectory(_directory);
        CreateWorkbook(
            Path.Combine(_directory, "ターゲット1900.xlsm"),
            ["作業シート", "問題解答リスト", "講師用", "生徒用"],
            [1900]);
        var profile = CreateProfile() with { WorkbookNameKeyword = "ターゲット1900" };

        var result = WorkbookDiscoveryService.Discover(_directory, [profile]);

        Assert.Empty(result.Candidates);
        Assert.Equal(1, result.RegisteredWorkbookCount);
    }

    [Fact]
    public void Discover_LeavesUnknownSheetLayoutUnselected()
    {
        Directory.CreateDirectory(_directory);
        CreateWorkbook(
            Path.Combine(_directory, "独自形式.xlsm"),
            ["入力", "一覧", "印刷"],
            []);

        var result = WorkbookDiscoveryService.Discover(_directory, [CreateProfile()]);

        var candidate = Assert.Single(result.Candidates);
        Assert.False(candidate.CanRegister);
        Assert.Empty(candidate.MatchingProfileIds);
    }

    [Fact]
    public void Discover_RemovesDuplicateWorkbookNamesAcrossFolders()
    {
        var firstFolder = Path.Combine(_directory, "one-drive");
        var secondFolder = Path.Combine(_directory, "google-drive");
        Directory.CreateDirectory(firstFolder);
        Directory.CreateDirectory(secondFolder);
        const string workbookName = "同じ教材.xlsm";
        CreateWorkbook(
            Path.Combine(firstFolder, workbookName),
            ["作業シート", "問題解答リスト", "講師用", "生徒用"],
            [100]);
        CreateWorkbook(
            Path.Combine(secondFolder, workbookName),
            ["作業シート", "問題解答リスト", "講師用", "生徒用"],
            [100]);

        var result = WorkbookDiscoveryService.Discover(
            [firstFolder, secondFolder],
            [CreateProfile()]);

        Assert.Single(result.Candidates);
        Assert.StartsWith(firstFolder, result.Candidates[0].WorkbookPath);
    }

    [Theory]
    [InlineData(20, 74, "word-pair-list")]
    [InlineData(172, 62, "grammar-choice")]
    public void Discover_InfersFinishedLayoutFromQuestionAndAnswerColumnWidths(
        double questionColumnWidth,
        double answerColumnWidth,
        string expectedLayoutId)
    {
        Directory.CreateDirectory(_directory);
        CreateWorkbook(
            Path.Combine(_directory, "分類名を含まない教材.xlsm"),
            ["作業シート", "問題解答リスト", "講師用", "生徒用"],
            [1, 100],
            questionColumnWidth,
            answerColumnWidth);
        var wordProfile = CreateProfile() with
        {
            Id = "word",
            FinishedLayoutId = "word-pair-list",
            FinishedLayoutName = "英単語・英熟語（1問1答）"
        };
        var grammarProfile = CreateProfile() with
        {
            Id = "grammar",
            FinishedLayoutId = "grammar-choice",
            FinishedLayoutName = "英文法・語法（選択問題）",
            MaximumQuestionCount = 25
        };

        var result = WorkbookDiscoveryService.Discover(
            _directory,
            [wordProfile, grammarProfile]);

        var candidate = Assert.Single(result.Candidates);
        Assert.Equal(expectedLayoutId, candidate.SuggestedFormatId);
        Assert.True(candidate.RequiresReview);
        Assert.Contains("内容から", candidate.Message);
    }

    [Fact]
    public void Discover_DoesNotClassifyWorkbookInJapaneseFolderAsEnglishVocabulary()
    {
        var japaneseFolder = Path.Combine(_directory, "テスト作成用Excel元ファイル(国語)");
        Directory.CreateDirectory(japaneseFolder);
        CreateWorkbook(
            Path.Combine(japaneseFolder, "ステップアップノート30.xlsm"),
            ["作業シート", "問題解答リスト", "講師用", "生徒用"],
            [1, 100],
            11,
            115);
        var wordProfile = CreateProfile() with
        {
            Id = "word",
            FinishedLayoutId = "word-pair-list",
            FinishedLayoutName = "英単語・英熟語（1問1答）"
        };
        var japaneseProfile = CreateProfile() with
        {
            Id = "japanese",
            FinishedLayoutId = "japanese-vocabulary",
            FinishedLayoutName = "国語（語彙・漢字・古文の1問1答）",
            UseWideAnswerLayout = true,
            MaximumQuestionCount = 50
        };

        var result = WorkbookDiscoveryService.Discover(
            japaneseFolder,
            [wordProfile, japaneseProfile]);

        var candidate = Assert.Single(result.Candidates);
        Assert.Equal("japanese-vocabulary", candidate.SuggestedFormatId);
        Assert.NotEqual("word-pair-list", candidate.SuggestedFormatId);
    }

    public void Dispose()
    {
        if (Directory.Exists(_directory))
        {
            Directory.Delete(_directory, recursive: true);
        }
    }

    private static ExcelTemplateProfile CreateProfile() => new()
    {
        Id = "standard",
        DisplayName = "標準形式",
        WorkbookNameKeyword = "既存教材",
        MaterialNameKeywords = ["既存教材"],
        MaximumQuestionNumber = 1000,
        MaximumQuestionCount = 100,
        WorkingSheetName = "作業シート",
        RangeStartCell = "G1",
        RangeEndCell = "G2",
        QuestionListSheetName = "問題解答リスト",
        TeacherSheetName = "講師用",
        StudentSheetName = "生徒用"
    };

    private static void CreateWorkbook(
        string path,
        IReadOnlyList<string> sheetNames,
        IReadOnlyList<int> questionNumbers,
        double? questionColumnWidth = null,
        double? answerColumnWidth = null)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        var sheets = new StringBuilder();
        var relationships = new StringBuilder();
        for (var index = 0; index < sheetNames.Count; index++)
        {
            var relationshipId = $"rId{index + 1}";
            sheets.Append(FormattableString.Invariant(
                $"<sheet name=\"{sheetNames[index]}\" sheetId=\"{index + 1}\" r:id=\"{relationshipId}\"/>"));
            relationships.Append(FormattableString.Invariant(
                $"<Relationship Id=\"{relationshipId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{index + 1}.xml\"/>"));
        }

        WriteEntry(
            archive,
            "xl/workbook.xml",
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            $"<sheets>{sheets}</sheets></workbook>");
        WriteEntry(
            archive,
            "xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"{relationships}</Relationships>");

        for (var index = 0; index < sheetNames.Count; index++)
        {
            var columns = sheetNames[index] == "問題解答リスト" &&
                          questionColumnWidth is not null &&
                          answerColumnWidth is not null
                ? FormattableString.Invariant(
                    $"<cols><col min=\"3\" max=\"3\" width=\"{questionColumnWidth}\"/><col min=\"4\" max=\"4\" width=\"{answerColumnWidth}\"/></cols>")
                : string.Empty;
            var cells = sheetNames[index] == "問題解答リスト"
                ? string.Concat(questionNumbers.Select((number, row) =>
                    $"<row r=\"{row + 2}\"><c r=\"B{row + 2}\"><v>{number}</v></c></row>"))
                : string.Empty;
            WriteEntry(
                archive,
                $"xl/worksheets/sheet{index + 1}.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                $"{columns}<sheetData>{cells}</sheetData></worksheet>");
        }
    }

    private static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
