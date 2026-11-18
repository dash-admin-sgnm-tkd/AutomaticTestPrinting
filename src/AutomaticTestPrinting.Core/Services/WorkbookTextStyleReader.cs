using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace AutomaticTestPrinting.Core.Services;

public sealed record WorkbookTextStyle(int Start, int Length, bool Bold, bool Italic, bool Underline, string VerticalPosition);

public static class WorkbookTextStyleReader
{
    public static IReadOnlyDictionary<string, IReadOnlyList<WorkbookTextStyle>> Read(string path, string sheetName)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry("xl/sharedStrings.xml")?.Open();
        var result = new Dictionary<string, IReadOnlyList<WorkbookTextStyle>>(StringComparer.Ordinal);
        if (stream is null) return result;
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var stylesByIndex = new Dictionary<int, IReadOnlyList<WorkbookTextStyle>>();
        var index = 0;
        foreach (var item in XDocument.Load(stream).Root?.Elements(ns + "si") ?? [])
        {
            var text = new StringBuilder();
            var styles = new List<WorkbookTextStyle>();
            foreach (var child in item.Elements().Where(element => element.Name == ns + "t" || element.Name == ns + "r"))
            {
                var part = child.Name == ns + "t" ? child.Value : (string?)child.Element(ns + "t") ?? "";
                var format = child.Element(ns + "rPr");
                var style = new WorkbookTextStyle(text.Length + 1, part.Length, Enabled("b"), Enabled("i"),
                    format?.Element(ns + "u") is { } underline && (string?)underline.Attribute("val") != "none",
                    (string?)format?.Element(ns + "vertAlign")?.Attribute("val") ?? "");
                if (part.Length > 0 && (style.Bold || style.Italic || style.Underline || style.VerticalPosition.Length > 0)) styles.Add(style);
                text.Append(part);

                bool Enabled(string name) => format?.Element(ns + name) is { } element && (string?)element.Attribute("val") is not ("0" or "false");
            }
            if (styles.Count > 0) stylesByIndex.Add(index, styles);
            index++;
        }
        XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using var workbookStream = archive.GetEntry("xl/workbook.xml")!.Open();
        using var relStream = archive.GetEntry("xl/_rels/workbook.xml.rels")!.Open();
        var sheet = XDocument.Load(workbookStream).Descendants(ns + "sheet").Single(s => (string?)s.Attribute("name") == sheetName);
        var id = (string?)sheet.Attribute(relNs + "id");
        var target = (string)XDocument.Load(relStream).Root!.Elements().Single(r => (string?)r.Attribute("Id") == id).Attribute("Target")!;
        target = target.StartsWith('/') ? target.TrimStart('/') : "xl/" + target;
        using var sheetStream = archive.GetEntry(target)!.Open();
        foreach (var cell in XDocument.Load(sheetStream).Descendants(ns + "c"))
            if ((string?)cell.Attribute("t") == "s" && int.TryParse((string?)cell.Element(ns + "v"), out var sharedIndex) && stylesByIndex.TryGetValue(sharedIndex, out var styles))
                result[(string)cell.Attribute("r")!] = styles;
        return result;
    }
}
