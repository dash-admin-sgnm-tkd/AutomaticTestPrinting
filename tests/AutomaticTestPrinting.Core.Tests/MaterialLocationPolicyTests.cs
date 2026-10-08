using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class MaterialLocationPolicyTests
{
    [Fact]
    public void SubjectOriginalWinsAndKunitachiIsExcluded()
    {
        string original = "C:/教材/①英語/単語/教材.xlsm";
        Assert.Equal([original], MaterialLocationPolicy.SelectCandidates([
            "C:/教材/国立校 確認テスト作成/①英語/教材.xlsm", "C:/教材/その他/教材.xlsm", original]));
    }

    [Fact]
    public void DoesNotChooseBetweenSubjectCopies()
    {
        Assert.Equal(2, MaterialLocationPolicy.SelectCandidates([
            "C:/①英語/a/教材.pdf", "C:/①英語/b/教材.pdf", "C:/国立校 確認テスト作成/教材.pdf"]).Length);
    }

    [Fact]
    public void AcademicKokuritsuLevelAndFilenameAreNotSchoolFolders()
    {
        Assert.False(MaterialLocationPolicy.IsExcluded("C:/①英語/国立/国立校対策.pdf"));
        Assert.False(MaterialLocationPolicy.IsPreferred("C:/その他/①英語教材.pdf"));
        Assert.Empty(MaterialLocationPolicy.SelectCandidates(["C:/国立校 確認テスト作成 - コピー/教材.pdf"]));
    }

    [Fact]
    public void StageFolderAndOverlappingRootsAreSupported()
    {
        string stage = "G:\\マイドライブ\\⓪段階突破テスト\\教材.pdf";
        Assert.Equal([stage], MaterialLocationPolicy.SelectCandidates([stage, stage, "G:/その他/教材.pdf"]));
    }

    [Fact]
    public void PdfResolverPrefersSubjectFileWithoutReadingExcludedCopy()
    {
        string preferred = "C:/教材/④古文/教材.pdf";
        Assert.Equal(preferred, PdfFileResolutionService.Resolve("教材.pdf", [
            "C:/教材/国立校 確認テスト作成/教材.pdf", preferred]));
        Assert.Throws<FileNotFoundException>(() => PdfFileResolutionService.Resolve("教材.pdf", [
            "C:/教材/国立校 確認テスト作成/教材.pdf"]));
    }
}
