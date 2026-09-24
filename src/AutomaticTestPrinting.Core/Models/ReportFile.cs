namespace AutomaticTestPrinting.Core.Models;

public sealed record ReportFile(string FullPath)
{
    public string FileName => Path.GetFileName(FullPath);
}
