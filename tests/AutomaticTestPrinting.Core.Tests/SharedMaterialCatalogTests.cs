using System.Text.Json;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class SharedMaterialCatalogTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), $"SharedCatalog-{Guid.NewGuid():N}");

    [Fact]
    public void BundledCatalog_ContainsAllThreeKindsAndOnlyReviewedPdfMappings()
    {
        var catalog = SharedMaterialCatalog.Load();
        var original = catalog.Items.Where(e => e.SourceSheet is not ("追加段階突破フォルダー" or "Excel")).ToArray();
        Assert.Equal(1, catalog.SchemaVersion);
        Assert.Contains(catalog.Items, e => e.Kind == "Excel");
        Assert.Contains(catalog.Items, e => e.Kind == "PDF");
        Assert.Contains(catalog.Items, e => e.Kind == "段階突破");
        Assert.Equal("1zyqrPPx79dzabrmqYBpyu9dhYzB1QEmuRyZI7_Mfl3E", catalog.SourceSpreadsheetId);
        Assert.Equal("12l-B7SZw_QKWVJvo06Y_UIfYqrrJVtsd", catalog.SourceFolderId);
        Assert.Equal(92, original.Count(e => e.Kind == "Excel"));
        Assert.Equal(97, catalog.Items.Count(e => e.Kind == "PDF"));
        Assert.Equal(44, original.Count(e => e.Kind == "段階突破"));
        Assert.Equal(9, catalog.Items.Count(e => e.Kind == "PDF" && e.AutoEnabled));
        Assert.Equal(44, original.Count(e => e.Kind == "段階突破" && e.AutoEnabled));
        Assert.Equal(original.Length, original.Select(e => e.SourceFileId).Distinct().Count());
        Assert.All(original, e => {
            Assert.False(string.IsNullOrWhiteSpace(e.SourceFileId));
            Assert.EndsWith(Assert.Single(e.FileNames), e.SourceRelativePath);
        });
        var unverified = catalog.Items.Single(e => e.FileNames.Contains("GMARCH&関関同立の古文_20150827本部.pdf"));
        Assert.False(unverified.AutoEnabled);
        Assert.Empty(unverified.Questions);
        Assert.DoesNotContain(original, e => e.DisplayName == "段階突破 英語 2019年度 日大 第1回");
        Assert.Contains(catalog.Items, e => e.DisplayName == "段階突破 漢文 2018年度 日大 第1回");
        Assert.DoesNotContain(catalog.Items, e => e.Kind == "Excel" && e.AutoEnabled);
        Assert.DoesNotContain(catalog.Items, e => e.FileNames.Contains("速読英単語2 上級編_20210818本部-LAPTOP-BQO04MKK.xlsm"));
        Assert.DoesNotContain(catalog.Items, e => e.FileNames.Contains("入門問題精講数学IAIIB_対照表.xlsx"));
        foreach (var name in new[]
        {
            "関正生のThe Rules 英語長文問題集1",
            "関正生のThe Rules 英語長文問題集3",
            "関正生の英語長文ポラリス1 標準レベル",
            "現代文読解力の開発講座"
        })
        {
            var entry = Assert.Single(catalog.Items, e => e.DisplayName == name);
            Assert.True(entry.AutoEnabled);
            Assert.NotEmpty(entry.Questions);
        }
        PdfMaterialProfileStore.Validate(SharedMaterialCatalog.CreatePdfProfiles());
    }

    [Fact]
    public void AdditionalExcelCatalog_PreservesPdfAndShowsAllSourceRowsWithoutAutomaticRegistration()
    {
        var catalog = SharedMaterialCatalog.Load();
        var added = catalog.Items.Where(e => e.SourceSheet == "Excel").ToArray();
        Assert.Equal(84, added.Length);
        Assert.Equal(176, catalog.Items.Count(e => e.Kind == "Excel"));
        Assert.Equal(10, added.Count(e => !string.IsNullOrWhiteSpace(e.RegistrationReviewReason)));
        Assert.Equal(103, catalog.Items.Count(e => e.SourceSheet == "Excel" || e.Notes.Contains(
            "追加調査情報：\n出典：1LU_JYYD86KUfWdS0jary4iX-aT76OYBqk6dvhpwVqzo / Excel ", StringComparison.Ordinal)));
        Assert.All(added, e =>
        {
            Assert.Equal("1LU_JYYD86KUfWdS0jary4iX-aT76OYBqk6dvhpwVqzo", e.SourceSpreadsheetId);
            Assert.False(e.AutoEnabled);
            Assert.Empty(e.Questions);
        });
        foreach (var row in new[] { 29, 30, 37, 74 })
        {
            var review = Assert.Single(added, e => e.SourceRow == row);
            Assert.Empty(review.FileNames);
            Assert.Contains("情報要確認", review.Status);
        }
        var derivative = Assert.Single(added, e => e.SourceRow == 4);
        Assert.Contains("派生語", derivative.DisplayName);
        Assert.Empty(derivative.Aliases);
        Assert.Equal("2026-10-08", catalog.UpdatedAt);
    }

    [Fact]
    public void MergeExcelCatalog_MatchesOnlyExactFileAndKeepsExistingIdentityAndSettings()
    {
        var original = new SharedMaterialEntry { Id = "existing", Kind = "Excel", DisplayName = "既存教材",
            FileNames = ["book.xlsm"], Notes = "元の情報", SourceFileId = "drive-id", SourceSheet = "元シート" };
        var pdf = new SharedMaterialEntry { Id = "pdf", Kind = "PDF", DisplayName = "PDF教材", FileNames = ["book.pdf"] };
        var primary = new SharedMaterialDatabase { SchemaVersion = 1, UpdatedAt = "2026-10-04", Items = [original, pdf] };
        var incoming = original with { Id = "new", DisplayName = "調査名", Notes = "調査情報", SourceSheet = "Excel" };
        var additional = new SharedMaterialDatabase { SchemaVersion = 1, UpdatedAt = "2026-10-08", Items = [incoming] };
        var merged = SharedMaterialCatalog.MergeExcelCatalog(primary, additional);
        Assert.Equal(2, merged.Items.Count);
        var result = Assert.Single(merged.Items, e => e.Kind == "Excel");
        Assert.Equal(original.Id, result.Id);
        Assert.Equal(original.DisplayName, result.DisplayName);
        Assert.Equal(original.SourceFileId, result.SourceFileId);
        Assert.Equal(original.SourceSheet, result.SourceSheet);
        Assert.Equal(original.Aliases, result.Aliases);
        Assert.Contains("調査情報", result.Notes);
        Assert.Equal(JsonSerializer.Serialize(pdf), JsonSerializer.Serialize(merged.Items.Single(e => e.Kind == "PDF")));
        var again = SharedMaterialCatalog.MergeExcelCatalog(merged, additional);
        Assert.Equal(result.Notes, again.Items.Single(e => e.Kind == "Excel").Notes);
        Assert.Throws<InvalidDataException>(() => SharedMaterialCatalog.MergeExcelCatalog(primary,
            additional with { Items = [incoming with { Kind = "PDF" }] }));
    }

    [Fact]
    public void MergeExcelCatalog_DoesNotMergeByMaterialNameOrGuessVersion()
    {
        var original = new SharedMaterialEntry { Id = "old", Kind = "Excel", DisplayName = "単語帳",
            FileNames = ["old.xlsm"] };
        var incoming = original with { Id = "new", FileNames = ["new.xlsm"] };
        var primary = new SharedMaterialDatabase { SchemaVersion = 1, Items = [original] };
        var additional = primary with { Items = [incoming] };
        Assert.Equal(2, SharedMaterialCatalog.MergeExcelCatalog(primary, additional).Items.Count);
        var uncertain = incoming with { Id = "review", FileNames = [], RegistrationReviewReason = "教材名の食い違い" };
        Assert.Equal(2, SharedMaterialCatalog.MergeExcelCatalog(primary, additional with { Items = [uncertain] }).Items.Count);
    }

    [Fact]
    public void AdditionalStageCatalog_IncludesMissingJapaneseTestsWithoutReplacingOriginalCatalog()
    {
        var catalog = SharedMaterialCatalog.Load();
        var additional = catalog.Items.Where(e => e.SourceSheet == "追加段階突破フォルダー").ToArray();
        Assert.NotEmpty(additional);
        Assert.All(additional, entry =>
        {
            Assert.Equal("段階突破", entry.Kind);
            Assert.True(entry.AutoEnabled);
            Assert.True(Assert.Single(entry.Questions).ExpectedPageCount > 0);
            Assert.Contains("SHA256", entry.Notes);
            Assert.EndsWith(Assert.Single(entry.FileNames), entry.SourceRelativePath);
        });
        foreach (var subject in new[] { "現代文", "古文" })
            Assert.Contains(additional, e => e.DisplayName == $"段階突破 {subject} 2020年度 日大 第2回");
        Assert.Equal(44, catalog.Items.Count(e => e.Kind == "段階突破" && e.SourceSheet != "追加段階突破フォルダー"));
    }

    [Fact]
    public void Parse_RejectsUnknownSchemaDuplicateIdsPathsAndAutomaticExcel()
    {
        var entry = new SharedMaterialEntry { Id = "a", Kind = "Excel", DisplayName = "教材", FileNames = ["book.xlsm"] };
        string Encode(params SharedMaterialEntry[] items) => JsonSerializer.Serialize(new SharedMaterialDatabase { SchemaVersion = 1, Items = items });
        Assert.Throws<InvalidDataException>(() => SharedMaterialCatalog.Parse("{\"schemaVersion\":2,\"items\":[]}"));
        Assert.Throws<InvalidDataException>(() => SharedMaterialCatalog.Parse(Encode(entry, entry)));
        Assert.Throws<InvalidDataException>(() => SharedMaterialCatalog.Parse(Encode(entry with { FileNames = ["../book.xlsm"] })));
        Assert.Throws<InvalidDataException>(() => SharedMaterialCatalog.Parse(Encode(entry with { AutoEnabled = true })));
    }

    [Fact]
    public void MergeProfiles_UserRegistrationWinsWithoutEditingBundledData()
    {
        var shared = SharedMaterialCatalog.CreatePdfProfiles();
        var original = shared.First(p => !p.IsStageTest);
        var user = original with { Id = "custom", Questions = [new(1, "custom.pdf", "1")] };
        var merged = PdfMaterialProfileStore.MergeProfiles(shared, [user]);
        Assert.Contains(user, merged);
        Assert.DoesNotContain(merged, p => p.Id == original.Id);
        Assert.Contains(original, SharedMaterialCatalog.CreatePdfProfiles());
    }

    [Fact]
    public void ExplicitPathLoad_DoesNotMixInSharedProfilesOrLoseOldJson()
    {
        Directory.CreateDirectory(_folder);
        var path = Path.Combine(_folder, "user.json");
        var user = new PdfMaterialProfile("user", "独自教材", [], [new(1, "user.pdf", "")]);
        PdfMaterialProfileStore.Save([user], path);
        Assert.Equal(user.Id, Assert.Single(PdfMaterialProfileStore.Load(path)).Id);
        Assert.Empty(PdfMaterialProfileStore.Load(Path.Combine(_folder, "missing.json")));
    }

    [Fact]
    public void OutOfScopeStageIsUnregisteredAndSharedPagesAreReviewRequired()
    {
        var ambiguous = PdfMaterialCatalog.Prepare(new("段階突破 漢文 2099年度 早稲田 第2回", "1-1", 1, 1, true), [_folder]);
        Assert.True(ambiguous.IsSupported); // Classified as stage PDF; still not registered or creatable.
        Assert.False(ambiguous.IsValid);
        Assert.Contains("未登録", ambiguous.Message);
        var sharedPages = PdfMaterialCatalog.Prepare(new("関正生のThe Rules 英語長文問題集1", "1-2", 1, 1, true), [_folder]);
        Assert.True(sharedPages.IsSupported);
        Assert.False(sharedPages.IsValid);
    }

    [Fact]
    public void PendingExcelCatalogEntry_DoesNotBecomeAutomaticProfile()
    {
        var result = ExcelTemplateCatalog.Prepare(new("高校英文読解をひとつひとつわかりやすく（改訂版） / 基本演習", "1-2", 1, 1, true), [_folder]);
        Assert.False(result.IsSupported);
        Assert.False(result.IsValid);
        Assert.Contains("カタログ", result.Message);
    }

    [Fact]
    public void StageRequest_UsesCompleteIdentityAndOutputsWholeFileWithoutSeparateAnswers()
    {
        var profile = SharedMaterialCatalog.CreatePdfProfiles().Single(p => p.DisplayName == "段階突破 漢文 2019年度 日大 第1回");
        CreatePdf(profile.Questions[0].FileName, 3);
        var ready = PdfMaterialCatalog.Prepare(new("段階突破 漢文 2019年度 日大レベル 第1回", "1-1", 1, 1, true), [_folder], [profile]);
        Assert.True(ready.IsValid, ready.Message);
        var result = PdfMaterialCatalog.Generate(new("A", profile.DisplayName, "1-1", 1, "report.pdf", 1, PdfQuestions: ready.Questions), _folder);
        Assert.Null(result.AnswerPdfPath);
        using var pdf = PdfReader.Open(result.ProblemPdfPath, PdfDocumentOpenMode.Import);
        Assert.Equal(3, pdf.PageCount);
        Assert.False(PdfMaterialCatalog.Prepare(new("段階突破 漢文 2020年度 日大 第1回", "1-1", 1, 1, true), [_folder], [profile]).IsValid);
        Assert.False(PdfMaterialCatalog.Prepare(new(profile.DisplayName, "1-2", 1, 1, true), [_folder], [profile]).IsValid);
    }

    [Fact]
    public void MappedPdf_RejectsDifferentPageCountAndRechecksBeforeGeneration()
    {
        var profile = new PdfMaterialProfile("book", "教材", [], [new(1, "book.pdf", "2-3", ExpectedPageCount: 3)]);
        CreatePdf("book.pdf", 2);
        var request = new NormalTestRequestCandidate("教材", "1-1", 1, 1, true);
        var bad = PdfMaterialCatalog.Prepare(request, [_folder], [profile]);
        Assert.False(bad.IsValid);
        Assert.Contains("総ページ数", bad.Message);
        CreatePdf("book.pdf", 3);
        var ready = PdfMaterialCatalog.Prepare(request, [_folder], [profile]);
        Assert.True(ready.IsValid, ready.Message);
        CreatePdf("book.pdf", 2);
        Assert.Throws<InvalidOperationException>(() => PdfMaterialCatalog.Generate(
            new("A", "教材", "1-1", 1, "report.pdf", 1, PdfQuestions: ready.Questions), _folder));
    }

    [Fact]
    public void StageProfile_RejectsPartialPagesOrSeparateAnswerOutput()
    {
        Assert.Throws<InvalidDataException>(() => PdfMaterialProfileStore.Validate([
            new("stage", "段階突破", [], [new(1, "stage.pdf", "1")], IsStageTest: true)]));
        Assert.Throws<InvalidDataException>(() => PdfMaterialProfileStore.Validate([
            new("stage", "段階突破", [], [new(1, "stage.pdf", "", "answers.pdf", "")], IsStageTest: true)]));
    }

    private void CreatePdf(string name, int pages)
    {
        Directory.CreateDirectory(_folder);
        using var document = new PdfDocument();
        for (var i = 0; i < pages; i++) document.AddPage();
        document.Save(Path.Combine(_folder, name));
    }

    public void Dispose()
    {
        if (Directory.Exists(_folder)) Directory.Delete(_folder, recursive: true);
    }
}
