using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfMaterialCatalogTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-PdfCatalog-{Guid.NewGuid():N}");

    [Theory]
    [InlineData("現代文")]
    [InlineData("古文")]
    [InlineData("漢文")]
    public void Prepare_StageSubjectFromReportMatchesCatalogButNotDifferentYearLevelOrAttempt(string subject)
    {
        CreatePdf("stage.pdf", 3);
        var profile = new PdfMaterialProfile("stage", $"段階突破 {subject} 2020年度 日大 第2回", [],
            [new(1, "stage.pdf", "")], IsStageTest: true);
        var name = $"段階突破 国語（{subject}） 2020年度 日大 第2回";
        var ready = PdfMaterialCatalog.Prepare(new(name, "1-1", 1, 1, true), [_folder], [profile]);
        Assert.True(ready.IsValid, ready.Message);
        Assert.Contains("全ページ", ready.Message);
        var result = PdfMaterialCatalog.Generate(new("A", name, "1-1", 1, "report.pdf", 1,
            PdfQuestions: ready.Questions), _folder);
        using var pdf = PdfReader.Open(result.ProblemPdfPath, PdfDocumentOpenMode.Import);
        Assert.Equal(3, pdf.PageCount);
        Assert.Null(result.AnswerPdfPath);
        foreach (var wrongName in new[] { name.Replace("2020", "2021", StringComparison.Ordinal),
            name.Replace("日大", "MARCH", StringComparison.Ordinal), name.Replace("第2回", "第1回", StringComparison.Ordinal) })
        {
            var wrong = PdfMaterialCatalog.Prepare(new(wrongName, "1-1", 1, 1, true), [_folder], [profile]);
            Assert.False(wrong.IsValid);
            Assert.True(wrong.IsSupported);
            Assert.Contains("段階突破PDFは未登録", wrong.Message);
        }
    }

    [Fact]
    public void Prepare_UnregisteredStageDoesNotFallBackToExcel()
    {
        var result = PdfMaterialCatalog.Prepare(new("段階突破 国語(現代文) 2020年度 日大 第2回", "1-1", 1, 1, true), [], []);
        Assert.True(result.IsSupported);
        Assert.False(result.IsValid);
        Assert.Contains("PDF教材を登録", result.Message);
    }

    [Fact]
    public void Prepare_UsesExactNormalizedAliasesAndFindsFilesRecursively()
    {
        var source = CreatePdf("book.pdf", 3);
        var profile = Profile([new(1, "book.pdf", "1"), new(2, "book.pdf", "2-3")]);
        var ready = PdfMaterialCatalog.Prepare(new("長 文 教 材（第１版）", "1-2", 1, 1, true), [_folder], [profile]);
        Assert.True(ready.IsValid, ready.Message);
        Assert.Equal(source, ready.Questions![0].Path);
        var wrong = PdfMaterialCatalog.Prepare(new("長文教材第1版派生", "1-2", 1, 1, true), [_folder], [profile]);
        Assert.False(wrong.IsSupported);
    }

    [Fact]
    public void Prepare_AllowsByteIdenticalPdfCopiesAcrossFolders()
    {
        var source = CreatePdf("copies.pdf", 2);
        var copyFolder = Path.Combine(_folder, "copy");
        Directory.CreateDirectory(copyFolder);
        File.Copy(source, Path.Combine(copyFolder, "copies.pdf"));
        var profile = Profile([new(1, "copies.pdf", "")]);
        var result = PdfMaterialCatalog.Prepare(new("長文教材第1版", "1-1", 1, 1, true), [_folder], [profile]);
        Assert.True(result.IsValid, result.Message);
        Assert.Single(result.Questions!);
    }

    [Fact]
    public void Prepare_RejectsDuplicateFileNamesEvenAcrossDifferentRoots()
    {
        CreatePdf("book.pdf", 1);
        CreatePdf("book.pdf", 1, "duplicate");
        var ready = PdfMaterialCatalog.Prepare(new("長文教材第1版", "1-1", 1, 1, true), [_folder], [Profile([new(1, "book.pdf", "")])]);
        Assert.False(ready.IsValid);
        Assert.Contains("同名PDF", ready.Message);
    }

    [Theory]
    [InlineData("1-3", 1)]
    [InlineData("1-2", 3)]
    [InlineData("0-1", 1)]
    public void Prepare_RejectsMissingMappingsOrInvalidCounts(string range, int count)
    {
        CreatePdf("book.pdf", 2);
        var ready = PdfMaterialCatalog.Prepare(new("長文教材第1版", range, count, 1, true), [_folder],
            [Profile([new(1, "book.pdf", "1"), new(2, "book.pdf", "2")])]);
        Assert.False(ready.IsValid);
    }

    [Fact]
    public void SaveAndLoad_PersistsMappingCreatesBackupAndRejectsDuplicateAliases()
    {
        Directory.CreateDirectory(_folder);
        var path = Path.Combine(_folder, "catalog.json");
        var profile = Profile([new(1, "book.pdf", "1")]);
        PdfMaterialProfileStore.Save([profile], path);
        Assert.Equal("1", Assert.Single(PdfMaterialProfileStore.Load(path)).Questions[0].Pages);
        PdfMaterialProfileStore.Save([profile], path);
        Assert.True(File.Exists(path + ".backup"));
        Assert.Throws<InvalidDataException>(() => PdfMaterialProfileStore.Save([profile, profile with { Id = "other" }], path));
        Assert.Single(PdfMaterialProfileStore.Load(path));
    }

    [Fact]
    public void Validate_RejectsDuplicateNumbersPathsAndPartialAnswers()
    {
        Assert.Throws<InvalidDataException>(() => PdfMaterialProfileStore.Validate([Profile([new(1, "book.pdf", "1"), new(1, "book.pdf", "2")])]));
        Assert.Throws<InvalidDataException>(() => PdfMaterialProfileStore.Validate([Profile([new(1, "../book.pdf", "1")])]));
        Assert.Throws<InvalidDataException>(() => PdfMaterialProfileStore.Validate([Profile([new(1, "book.pdf", "1", "answers.pdf", "1"), new(2, "book.pdf", "2")])]));
    }

    [Fact]
    public void Generate_CombinesMappedPagesAndCorrespondingAnswersWithoutModifyingSource()
    {
        var source = CreatePdf("book.pdf", 4);
        var bytes = File.ReadAllBytes(source);
        var ready = PdfMaterialCatalog.Prepare(new("長文教材第1版", "1-2", 2, 1, true), [_folder],
            [Profile([new(1, "book.pdf", "1", "book.pdf", "3"), new(2, "book.pdf", "2", "book.pdf", "4")])]);
        var plan = new TestGenerationPlan("生徒A", "長文教材第1版", "1-2", 2, "report.pdf", 1, PdfQuestions: ready.Questions);
        var result = PdfMaterialCatalog.Generate(plan, _folder);
        using var problem = PdfReader.Open(result.ProblemPdfPath, PdfDocumentOpenMode.Import);
        using var answer = PdfReader.Open(result.AnswerPdfPath!, PdfDocumentOpenMode.Import);
        Assert.Equal(2, problem.PageCount);
        Assert.Equal(2, answer.PageCount);
        Assert.Equal(1, problem.Pages[0].Rotate / 90); // unique marker in each source page
        Assert.Equal(3, answer.Pages[0].Rotate / 90);
        Assert.Equal(bytes, File.ReadAllBytes(source));
        Assert.Contains("1,2", File.ReadAllText(Path.Combine(Path.GetDirectoryName(result.ProblemPdfPath)!, "作成内容.txt")));
    }

    [Fact]
    public void Generate_WholeFileWithoutAnswersCanBeMergedWithExcelOutput()
    {
        var source = CreatePdf("stage.pdf", 1);
        var result = PdfMaterialCatalog.Generate(new("A", "段階突破", "1-1", 1, "report.pdf", 1,
            PdfQuestions: [new(1, source, "", null, "")]), _folder);
        Assert.Null(result.AnswerPdfPath);
        var merged = PdfBatchMergeService.Merge([result, new(source, source)], _folder, DateTimeOffset.Now, duplex: true);
        using var problem = PdfReader.Open(merged.ProblemPdfPath, PdfDocumentOpenMode.Import);
        using var answer = PdfReader.Open(merged.AnswerPdfPath!, PdfDocumentOpenMode.Import);
        Assert.Equal(3, problem.PageCount);
        Assert.Single(answer.Pages);
        var problemOnly = PdfBatchMergeService.Merge([result], _folder, DateTimeOffset.Now);
        Assert.Null(problemOnly.AnswerPdfPath);
    }

    [Fact]
    public void Generate_ReportsInvalidMappedPageWithQuestionAndSource()
    {
        var source = CreatePdf("book.pdf", 1);
        var exception = Assert.Throws<InvalidOperationException>(() => PdfMaterialCatalog.Generate(
            new("A", "教材", "7-7", 1, "report.pdf", 1, PdfQuestions: [new(7, source, "2", null, "")]), _folder));
        Assert.Contains("問題番号 7", exception.Message);
        Assert.Contains(source, exception.Message);
    }

    private static PdfMaterialProfile Profile(IReadOnlyList<PdfQuestionMapping> questions) => new("book", "長文教材第1版", ["長文教材"], questions);

    private string CreatePdf(string name, int pages, string subfolder = "source")
    {
        var folder = Path.Combine(_folder, subfolder);
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, name);
        using var document = new PdfDocument();
        for (var index = 1; index <= pages; index++) document.AddPage().Rotate = (index % 4) * 90;
        document.Save(path);
        return path;
    }

    public void Dispose()
    {
        if (Directory.Exists(_folder)) Directory.Delete(_folder, recursive: true);
    }
}
