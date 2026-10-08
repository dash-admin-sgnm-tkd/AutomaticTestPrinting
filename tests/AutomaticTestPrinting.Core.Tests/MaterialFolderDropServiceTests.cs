using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class MaterialFolderDropServiceTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "FolderDropTests-" + Guid.NewGuid().ToString("N"));

    private string Folder(string name)
    {
        var path = Path.Combine(_root, name);
        Directory.CreateDirectory(path);
        return path;
    }

    [Fact]
    public void Plan_AddsMultipleFoldersInDropOrderWithoutDuplicates()
    {
        var a = Folder("a"); var b = Folder("b"); var c = Folder("c");
        var result = MaterialFolderDropService.Plan([a, b, b + Path.DirectorySeparatorChar, c], [a]);
        Assert.Equal(new[] { b, c }, result.AddedFolders);
        Assert.Equal(2, result.DuplicateCount);
        Assert.Equal(0, result.IgnoredCount);
    }

    [Fact]
    public void Plan_IgnoresFilesAndMissingPathsWithoutChangingThem()
    {
        var folder = Folder("教材");
        var file = Path.Combine(folder, "教材.xlsm");
        File.WriteAllText(file, "untouched");
        var result = MaterialFolderDropService.Plan([file, Path.Combine(_root, "missing"), "", folder], []);
        Assert.Equal(new[] { folder }, result.AddedFolders);
        Assert.Equal(3, result.IgnoredCount);
        Assert.Equal("untouched", File.ReadAllText(file));
    }

    [Fact]
    public void Plan_NormalizesExistingTrailingSeparators()
    {
        var folder = Folder("a");
        var result = MaterialFolderDropService.Plan([folder], [folder + Path.DirectorySeparatorChar]);
        Assert.Empty(result.AddedFolders);
        Assert.Equal(1, result.DuplicateCount);
    }

    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }
}
