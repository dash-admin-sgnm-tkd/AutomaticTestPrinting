using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ConfirmedReportMaterialNamesTests
{
    [Fact]
    public void SavedChoiceSurvivesReadAndDoesNotChangeDifferentReportNames()
    {
        var path = Path.GetTempFileName();
        File.WriteAllText(path, "{}");
        try
        {
            ConfirmedReportMaterialNames.Save("独自名2", "教材2", path);
            ConfirmedReportMaterialNames.Save("独自名3", "教材3", path);
            Assert.Equal("教材2", ConfirmedReportMaterialNames.Resolve("独自名2", path));
            Assert.Equal("教材3", ConfirmedReportMaterialNames.Resolve("独自名3", path));
            Assert.Equal("独自名2別版", ConfirmedReportMaterialNames.Resolve("独自名2別版", path));
        }
        finally { File.Delete(path); }
    }
}
