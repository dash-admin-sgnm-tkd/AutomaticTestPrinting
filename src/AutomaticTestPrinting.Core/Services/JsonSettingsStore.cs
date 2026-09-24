using System.Text.Json;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public sealed class JsonSettingsStore(string settingsPath)
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true
    };

    public async Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(settingsPath))
        {
            return new AppSettings();
        }

        try
        {
            await using var stream = File.OpenRead(settingsPath);
            return await JsonSerializer.DeserializeAsync<AppSettings>(
                       stream,
                       SerializerOptions,
                       cancellationToken)
                   ?? new AppSettings();
        }
        catch (JsonException)
        {
            return new AppSettings();
        }
        catch (IOException)
        {
            return new AppSettings();
        }
    }

    public async Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default)
    {
        var directory = Path.GetDirectoryName(settingsPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        await using var stream = File.Create(settingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, SerializerOptions, cancellationToken);
    }
}
