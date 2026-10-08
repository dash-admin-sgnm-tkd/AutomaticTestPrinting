using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class ExcelMaterialEditingService
{
    // Exact filenames take priority over display names; never infer editions from aliases.
    public static IReadOnlyList<ExcelTemplateProfile> FindRegistered(
        IReadOnlyList<ExcelTemplateProfile> profiles, IReadOnlyList<string> fileNames, string displayName)
    {
        var byFile = profiles.Where(p => fileNames.Any(name =>
            string.Equals(p.WorkbookFileName, name, StringComparison.OrdinalIgnoreCase) ||
            (string.IsNullOrWhiteSpace(p.WorkbookFileName) &&
             string.Equals(p.WorkbookNameKeyword, Path.GetFileNameWithoutExtension(name), StringComparison.OrdinalIgnoreCase))))
            .ToArray();
        return byFile.Length > 0 ? byFile : profiles.Where(p =>
            string.Equals(p.DisplayName, displayName, StringComparison.OrdinalIgnoreCase)).ToArray();
    }

    public static IReadOnlyList<ExcelTemplateProfile> Upsert(
        IReadOnlyList<ExcelTemplateProfile> userProfiles, IReadOnlyList<ExcelTemplateProfile> updates) =>
        userProfiles.Where(p => !updates.Any(u => string.Equals(u.Id, p.Id, StringComparison.OrdinalIgnoreCase)))
            .Concat(updates).ToArray();
}
