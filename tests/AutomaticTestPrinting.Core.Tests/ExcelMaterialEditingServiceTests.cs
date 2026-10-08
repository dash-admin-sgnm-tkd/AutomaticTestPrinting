using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ExcelMaterialEditingServiceTests
{
    [Fact]
    public void ExactFileWinsOverDifferentDisplayNameAndOtherEditions()
    {
        var original = Profile("built-in", "旧表示名", "six.xlsm");
        var otherEdition = Profile("old-edition", "カタログ表示名", "five.xlsm");
        Assert.Equal(original, Assert.Single(ExcelMaterialEditingService.FindRegistered(
            [otherEdition, original], ["six.xlsm"], "カタログ表示名")));
    }

    [Fact]
    public void LegacyExactStemAndExactDisplayNameCanFindRegistration()
    {
        var legacy = Profile("custom-old", "登録名", "book.xlsm") with { WorkbookFileName = "" };
        Assert.Equal(legacy, Assert.Single(ExcelMaterialEditingService.FindRegistered([legacy], ["book.xlsm"], "")));
        Assert.Equal(legacy, Assert.Single(ExcelMaterialEditingService.FindRegistered([legacy], ["renamed.xlsm"], "登録名")));
        Assert.Empty(ExcelMaterialEditingService.FindRegistered([legacy], ["book-new.xlsm"], "略称"));
    }

    [Fact]
    public void AmbiguousFileKeepsAllCandidatesRatherThanSilentlyEditingOne()
    {
        var first = Profile("one", "教材1", "book.xlsm");
        var second = Profile("two", "教材2", "book.xlsm");
        Assert.Equal(2, ExcelMaterialEditingService.FindRegistered([first, second], ["book.xlsm"], "教材1").Count);
    }

    [Fact]
    public void UpsertPreservesBuiltInOverridesAndUnrelatedCustomSettings()
    {
        var original = Profile("built-in", "教材", "book.xlsm");
        var unrelated = Profile("custom-other", "別教材", "other.xlsm");
        var edited = original with { DisplayName = "変更後", MaximumQuestionNumber = 2500 };
        var result = ExcelMaterialEditingService.Upsert([original, unrelated], [edited]);
        Assert.Equal(2, result.Count);
        Assert.Contains(edited, result);
        Assert.Contains(unrelated, result);
        Assert.Equal(1, result.Count(p => p.Id == original.Id));
        var next = ExcelMaterialEditingService.Upsert(result, [Profile("custom-new", "追加", "new.xlsm")]);
        Assert.Contains(edited, next);
        Assert.Equal(3, next.Count);
    }

    [Fact]
    public void UpsertBuiltInWithoutExistingOverrideDoesNotCopyOtherBuiltIns()
    {
        var edited = Profile("built-in", "変更後", "book.xlsm");
        Assert.Equal(edited, Assert.Single(ExcelMaterialEditingService.Upsert([], [edited])));
    }

    private static ExcelTemplateProfile Profile(string id, string name, string file) => new()
    {
        Id = id, DisplayName = name, WorkbookFileName = file,
        WorkbookNameKeyword = Path.GetFileNameWithoutExtension(file), MaterialNameKeywords = ["略称"],
        MaximumQuestionNumber = 1900, MaximumQuestionCount = 50
    };
}
