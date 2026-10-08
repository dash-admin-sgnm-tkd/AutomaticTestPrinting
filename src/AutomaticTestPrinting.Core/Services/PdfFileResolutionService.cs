using System.Security.Cryptography;

namespace AutomaticTestPrinting.Core.Services;

public static class PdfFileResolutionService
{
    public static string Resolve(string fileName, IReadOnlyCollection<string> files)
    {
        var found = MaterialLocationPolicy.SelectCandidates(files.Where(path =>
            string.Equals(Path.GetFileName(path), fileName, StringComparison.OrdinalIgnoreCase)));
        if (found.Length == 0) throw new FileNotFoundException($"参照対象の教材フォルダーにPDFがありません（国立校側は参照しません）：{fileName}");
        if (found.Length == 1) return found[0];
        // Multiple search roots can contain copies of the same PDF. Only byte-identical copies are interchangeable.
        var length = new FileInfo(found[0]).Length;
        if (found.Any(path => new FileInfo(path).Length != length)) throw Ambiguous(fileName, found);
        byte[] Hash(string path)
        {
            using var stream = File.OpenRead(path);
            return SHA256.HashData(stream);
        }
        var expected = Hash(found[0]);
        if (found.Skip(1).Any(path => !Hash(path).AsSpan().SequenceEqual(expected))) throw Ambiguous(fileName, found);
        return found[0];
    }

    private static InvalidOperationException Ambiguous(string fileName, string[] paths) =>
        new($"内容の異なる同名PDFが複数あります。使用する原本を確認してください：{fileName}\n{string.Join('\n', paths)}");
}
