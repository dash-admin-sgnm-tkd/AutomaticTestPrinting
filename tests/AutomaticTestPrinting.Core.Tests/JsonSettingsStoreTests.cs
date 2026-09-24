using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class JsonSettingsStoreTests
{
    [Fact]
    public async Task SaveAndLoadAsync_RoundTripsSettings()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "settings.json");

        try
        {
            var store = new JsonSettingsStore(path);
            var expected = new AppSettings
            {
                MaterialFolder = @"C:\OneDrive\教材",
                OutputFolder = @"C:\OneDrive\出力"
            };

            await store.SaveAsync(expected);
            var actual = await store.LoadAsync();

            Assert.Equal(expected.MaterialFolder, actual.MaterialFolder);
            Assert.Equal(expected.OutputFolder, actual.OutputFolder);
        }
        finally
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public async Task LoadAsync_ReturnsDefaultsForInvalidJson()
    {
        var path = Path.GetTempFileName();

        try
        {
            await File.WriteAllTextAsync(path, "not-json");
            var settings = await new JsonSettingsStore(path).LoadAsync();

            Assert.Null(settings.MaterialFolder);
            Assert.Null(settings.OutputFolder);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
