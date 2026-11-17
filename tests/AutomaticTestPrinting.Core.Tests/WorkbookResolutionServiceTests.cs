using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class WorkbookResolutionServiceTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "WorkbookChoiceTests-" + Guid.NewGuid().ToString("N"));
    private static ExcelTemplateProfile Profile => new() { Id = "test", WorkbookFileName = "教材.xlsm" };

    private string Add(string relative, string contents)
    {
        var path = Path.GetFullPath(Path.Combine(_root, relative));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, contents);
        return path;
    }

    [Fact]
    public void Resolve_PrefersSubjectFolderAndExcludesSchoolCopy()
    {
        var preferred = Add("①英語/教材.xlsm", "original");
        Add("国立校 確認テスト作成/①英語/教材.xlsm", "different school copy");
        Add("その他/教材.xlsm", "other copy");
        var result = WorkbookResolutionService.Resolve([_root], Profile);
        Assert.Equal(preferred, result.Path);
        Assert.Single(result.Candidates);
    }

    [Fact]
    public void Resolve_DoesNotUseRememberedSchoolCopyOrSilentlyChangeSelection()
    {
        var selected = Add("国立校 確認テスト作成/教材.xlsm", "school");
        var preferred = Add("①英語/教材.xlsm", "original");
        var choice = new WorkbookChoice(Path.Combine("国立校 確認テスト作成", "教材.xlsm"), WorkbookResolutionService.GetHash(selected));
        var result = WorkbookResolutionService.Resolve([_root], Profile, choice);
        Assert.Null(result.Path);
        Assert.Equal(preferred, Assert.Single(result.Candidates));
        Assert.Contains("参照対象外", result.Message);
    }

    [Fact]
    public void Resolve_ExcludesSchoolFolderEvenWhenItIsTheSearchRoot()
    {
        Add("国立校 確認テスト作成/教材.xlsm", "school");
        var result = WorkbookResolutionService.Resolve([Path.Combine(_root, "国立校 確認テスト作成")], Profile);
        Assert.Null(result.Path);
        Assert.Empty(result.Candidates);
        Assert.Contains("国立校", result.Message);
    }

    [Fact]
    public void Resolve_IdenticalCopiesAreAccepted()
    {
        Add("a/教材.xlsm", "same"); Add("b/教材.xlsm", "same");
        var result = WorkbookResolutionService.Resolve([_root], Profile);
        Assert.NotNull(result.Path);
        Assert.Equal(2, result.Candidates.Count);
        Assert.Contains("同一", result.Message);
    }

    [Fact]
    public void Resolve_DifferentCopiesRequireSelectionAndRememberSelectedCopy()
    {
        Add("a/教材.xlsm", "first");
        var selected = Add("b/教材.xlsm", "second");
        Assert.Null(WorkbookResolutionService.Resolve([_root], Profile).Path);
        var choice = new WorkbookChoice(Path.Combine("b", "教材.xlsm"), WorkbookResolutionService.GetHash(selected));
        Assert.Equal(selected, WorkbookResolutionService.Resolve([_root], Profile, choice).Path);
    }

    [Fact]
    public void Resolve_ChoiceSurvivesChangedRootAndFolderOrder()
    {
        var original = Add("old/b/教材.xlsm", "selected");
        var choice = new WorkbookChoice(Path.Combine("b", "教材.xlsm"), WorkbookResolutionService.GetHash(original));
        var moved = Add("new/b/教材.xlsm", "selected");
        Add("other/b/教材.xlsm", "different");
        Assert.Equal(moved, WorkbookResolutionService.Resolve([Path.Combine(_root, "other"), Path.Combine(_root, "new")], Profile, choice).Path);
    }

    [Fact]
    public void Resolve_ChangedOrMissingSelectionDoesNotFallBack()
    {
        var selected = Add("a/教材.xlsm", "original");
        Add("b/教材.xlsm", "other");
        var choice = new WorkbookChoice(Path.Combine("a", "教材.xlsm"), WorkbookResolutionService.GetHash(selected));
        File.WriteAllText(selected, "updated and longer");
        Assert.Null(WorkbookResolutionService.Resolve([_root], Profile, choice).Path);
        File.Delete(selected);
        Assert.Null(WorkbookResolutionService.Resolve([_root], Profile, choice).Path);
    }

    [Fact]
    public void Resolve_FollowsIdenticalWorkbookAfterMoveWithLegacyChoice()
    {
        var original = Add("a/教材.xlsm", "selected");
        var choice = new WorkbookChoice(Path.Combine("a", "教材.xlsm"), WorkbookResolutionService.GetHash(original));
        var moved = Add("b/教材.xlsm", "selected");
        File.Delete(original);
        Assert.Equal(moved, WorkbookResolutionService.Resolve([_root], Profile, choice).Path);
    }

    [Fact]
    public void Resolve_FollowsRenamedWorkbookOnlyWhenOneEligibleFileHasIdenticalContent()
    {
        var original = Add("a/教材.xlsm", "selected");
        var choice = new WorkbookChoice(Path.Combine("a", "教材.xlsm"),
            WorkbookResolutionService.GetHash(original), new FileInfo(original).Length);
        File.Delete(original);
        var renamed = Add("b/改名.xlsm", "selected");
        Add("b/別教材.xlsm", "different");
        Add("国立校 確認テスト作成/別名.xlsm", "selected");
        Assert.Equal(renamed, WorkbookResolutionService.Resolve([_root], Profile, choice).Path);

        Add("c/改名コピー.xlsm", "selected");
        Assert.Null(WorkbookResolutionService.Resolve([_root], Profile, choice).Path);
    }

    [Fact]
    public void Resolve_VerifiedContentDisambiguatesSameNameWithoutUsingWrongCopy()
    {
        var correct = Add("①英語/確認済み/教材.xlsm", "original");
        Add("①英語/別形式/教材.xlsm", "different");
        var profile = Profile with { VerifiedWorkbookSha256 = WorkbookResolutionService.GetHash(correct) };
        Assert.Equal(correct, WorkbookResolutionService.Resolve([_root], profile).Path);
        File.Delete(correct);
        var result = WorkbookResolutionService.Resolve([_root], profile);
        Assert.Null(result.Path);
        Assert.Empty(result.Candidates);
        Assert.Contains("内容が異なります", result.Message);
    }

    [Fact]
    public void Resolve_DoesNotAcceptAnotherEdition()
    {
        Add("教材旧版.xlsm", "other edition");
        Assert.Null(WorkbookResolutionService.Resolve([_root], Profile).Path);
    }

    [Fact]
    public void Resolve_KnownCandidatesSupportSavedSelectionWithoutDiscoveringOtherFiles()
    {
        var selected = Add("a/教材.xlsm", "selected");
        Add("b/教材.xlsm", "different");
        var choice = new WorkbookChoice(Path.Combine("a", "教材.xlsm"), WorkbookResolutionService.GetHash(selected));
        var result = WorkbookResolutionService.Resolve([_root], Profile, choice, [selected]);
        Assert.Equal(selected, result.Path);
        Assert.Single(result.Candidates);
    }

    [Fact]
    public void Resolve_KnownCandidatesStillRejectOutsideRootsWrongNamesAndMissingSelection()
    {
        var selected = Add("a/教材.xlsm", "selected");
        var otherEdition = Add("a/教材旧版.xlsm", "wrong");
        var outside = Add("outside/教材.xlsm", "wrong");
        var choice = new WorkbookChoice("教材.xlsm", WorkbookResolutionService.GetHash(selected));
        var result = WorkbookResolutionService.Resolve([Path.Combine(_root, "a")], Profile, choice, [selected, otherEdition, outside]);
        Assert.Equal(selected, result.Path);
        Assert.Single(result.Candidates);
        File.Delete(selected);
        Assert.Null(WorkbookResolutionService.Resolve([Path.Combine(_root, "a")], Profile, choice, [selected, otherEdition, outside]).Path);
    }

    [Fact]
    public void Store_RoundTripsAndPreservesOtherChoices()
    {
        var selected = Add("a/教材.xlsm", "data");
        var choice = new WorkbookChoice(Path.Combine("a", "教材.xlsm"), WorkbookResolutionService.GetHash(selected));
        var path = Path.Combine(_root, "choices.json");
        WorkbookChoiceStore.Save("first", choice, path);
        WorkbookChoiceStore.Save("second", choice, path);
        Assert.Equal(choice, WorkbookChoiceStore.Load(path)["first"]);
        Assert.Equal(2, WorkbookChoiceStore.Load(path).Count);
        Assert.DoesNotContain(_root, File.ReadAllText(path), StringComparison.Ordinal);
    }

    [Fact]
    public void Store_CorruptConfigurationIsNotOverwritten()
    {
        var path = Add("choices.json", "invalid JSON");
        Assert.Throws<System.Text.Json.JsonException>(() => WorkbookChoiceStore.Save("test", new("a", new string('A', 64)), path));
        Assert.Equal("invalid JSON", File.ReadAllText(path));
    }

    [Fact]
    public void RelativePath_RejectsOutsideRoot()
    {
        Assert.Null(WorkbookResolutionService.GetRelativePath(Path.Combine(_root, "a"), Path.Combine(_root, "b", "教材.xlsm")));
    }

    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }
}
