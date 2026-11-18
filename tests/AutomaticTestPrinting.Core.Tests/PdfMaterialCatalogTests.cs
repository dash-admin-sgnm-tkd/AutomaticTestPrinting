using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfMaterialCatalogTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-PdfCatalog-{Guid.NewGuid():N}");

    [Fact]
    public void Prepare_OnlyExplicitSparseProfilesSkipUnprintedItemNumbers()
    {
        CreatePdf("sparse.pdf", 2);
        var sparse = new PdfMaterialProfile("sparse", "練習問題", [],
            [new(3,"sparse.pdf","1"),new(5,"sparse.pdf","2"),new(6,"sparse.pdf","2")], AllowSparseRanges:true);
        var ready = PdfMaterialCatalog.Prepare(new("練習問題","1-6",3,1,true),[_folder],[sparse]);
        Assert.True(ready.IsValid,ready.Message);
        Assert.Equal([3,5,6],ready.Questions!.Select(q=>q.Number));
        Assert.Contains("未収録",ready.Message);
        var strict = PdfMaterialCatalog.Prepare(new("練習問題","3-6",3,1,true),[_folder],[sparse with{AllowSparseRanges=false}]);
        Assert.False(strict.IsValid);
        Assert.Contains("未登録",strict.Message);
        var empty = PdfMaterialCatalog.Prepare(new("練習問題","4-4",1,1,true),[_folder],[sparse]);
        Assert.False(empty.IsValid);
        Assert.Contains("収録された練習問題がありません",empty.Message);
        var beyond = PdfMaterialCatalog.Prepare(new("練習問題","3-7",1,1,true),[_folder],[sparse]);
        Assert.False(beyond.IsValid);
        var tooMany = PdfMaterialCatalog.Prepare(new("練習問題","1-6",4,1,true),[_folder],[sparse]);
        Assert.False(tooMany.IsValid);
        Assert.Contains("1-3問",tooMany.Message);
    }

    [Fact]
    public void SparseRangeSettingSurvivesSavingAndDoesNotApplyToStageTests()
    {
        Directory.CreateDirectory(_folder);
        var path=Path.Combine(_folder,"profiles.json");
        var profile=new PdfMaterialProfile("sparse","疎な教材",[],[new(3,"sparse.pdf","1")],AllowSparseRanges:true);
        PdfMaterialProfileStore.Save([profile],path);
        Assert.True(Assert.Single(PdfMaterialProfileStore.Load(path)).AllowSparseRanges);
        Assert.Throws<InvalidDataException>(()=>PdfMaterialProfileStore.Validate(
            [profile with{IsStageTest=true,Questions=[new(1,"stage.pdf","")]}]));
    }

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
    public void Prepare_ReusesCandidateListAndRejectsFilesRemovedSinceCapture()
    {
        var source = CreatePdf("book.pdf", 3);
        var request = new NormalTestRequestCandidate("長文教材第1版", "1-1", 1, 1, true);
        var profiles = new[] { Profile([new(1, "book.pdf", "1")]) };

        var initial = PdfMaterialCatalog.Prepare(request, [_folder], profiles, [source]);
        Assert.True(initial.IsValid, initial.Message);
        Assert.Equal(source, Assert.Single(initial.Questions!).Path);

        File.Delete(source);
        var removed = PdfMaterialCatalog.Prepare(request, [_folder], profiles, [source]);
        Assert.False(removed.IsValid);
        Assert.Contains("PDFがありません", removed.Message);
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
    public void Generate_CombinesSharedProblemAndAnswerPagesOnceWithoutModifyingSource()
    {
        var source = CreatePdf("book.pdf", 4);
        var bytes = File.ReadAllBytes(source);
        var ready = PdfMaterialCatalog.Prepare(new("長文教材第1版", "1-2", 2, 1, true), [_folder],
            [Profile([new(1, "book.pdf", "1-2", "book.pdf", "3"), new(2, "book.pdf", "2", "book.pdf", "3-4")])]);
        var plan = new TestGenerationPlan("生徒A", "長文教材第1版", "1-2", 2, "report.pdf", 1, PdfQuestions: ready.Questions);
        var result = PdfMaterialCatalog.Generate(plan, _folder);
        using var problem = PdfReader.Open(result.ProblemPdfPath, PdfDocumentOpenMode.Import);
        using var answer = PdfReader.Open(result.AnswerPdfPath!, PdfDocumentOpenMode.Import);
        Assert.Equal(2, problem.PageCount);
        Assert.Equal(2, answer.PageCount);
        Assert.Equal(90, problem.Pages[0].Rotate);
        Assert.Equal(180, problem.Pages[1].Rotate);
        Assert.Equal(270, answer.Pages[0].Rotate);
        Assert.Equal(0, answer.Pages[1].Rotate);
        Assert.Equal(bytes, File.ReadAllBytes(source));
        Assert.Equal(_folder, Path.GetDirectoryName(result.ProblemPdfPath));
        Assert.Equal(_folder, Path.GetDirectoryName(result.AnswerPdfPath));
        Assert.Empty(Directory.EnumerateFiles(_folder, "*.txt"));
        Assert.Single(Directory.EnumerateDirectories(_folder)); // source PDFs only

        var repeated = PdfMaterialCatalog.Generate(plan, _folder);
        Assert.NotEqual(result.ProblemPdfPath, repeated.ProblemPdfPath);
        Assert.NotEqual(result.AnswerPdfPath, repeated.AnswerPdfPath);
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
