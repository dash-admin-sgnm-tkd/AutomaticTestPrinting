using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfMaterialGenerationServiceTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-PdfMaterials-{Guid.NewGuid():N}");

    [Fact]
    public void ParsePages_PreservesRequestedOrderAndExpandsRanges()
    {
        Assert.Equal([5, 1, 2, 3], PdfMaterialGenerationService.ParsePages("5, 1-3", 5));
        Assert.Equal([1, 2, 3], PdfMaterialGenerationService.ParsePages("", 3));
    }

    [Theory]
    [InlineData("0")]
    [InlineData("4")]
    [InlineData("3-1")]
    [InlineData("1,1")]
    [InlineData("1-2,2-3")]
    [InlineData("1,")]
    [InlineData("1--2")]
    [InlineData("one")]
    public void ParsePages_RejectsInvalidOrDuplicatePages(string pages)
    {
        Assert.Throws<ArgumentException>(() => PdfMaterialGenerationService.ParsePages(pages, 3));
    }

    [Fact]
    public void Generate_ExtractsCorrectPagesAndKeepsSourceUnchanged()
    {
        var source = CreatePdf("source.pdf", 3);
        var original = File.ReadAllBytes(source);
        var request = new PdfMaterialRequest("生徒", "長文", source, "3,1", false);
        var result = PdfMaterialGenerationService.Generate([request], _folder, duplex: true);

        Assert.Empty(result.Errors);
        var output = Assert.Single(result.Outputs);
        Assert.Equal(2, output.PageCount);
        using var pdf = PdfReader.Open(output.Path, PdfDocumentOpenMode.Import);
        Assert.Equal(303, pdf.Pages[0].Width.Point, precision: 2);
        Assert.Equal(301, pdf.Pages[1].Width.Point, precision: 2);
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.NotNull(result.ProblemPdfPath);
        Assert.Null(result.AnswerPdfPath);
    }

    [Fact]
    public void Generate_SeparatesAnswersAndPadsDuplexBoundaries()
    {
        var source = CreatePdf("source.pdf", 3);
        var result = PdfMaterialGenerationService.Generate(
            [new("A", "長文", source, "1", false), new("B", "長文", source, "2-3", false),
             new("A", "解答", source, "3", true)], _folder, duplex: true);
        using var problem = PdfReader.Open(result.ProblemPdfPath!, PdfDocumentOpenMode.Import);
        Assert.Equal(4, problem.PageCount);
        Assert.Equal(301, problem.Pages[1].Width.Point, precision: 2); // blank matching preceding page
        Assert.Equal(302, problem.Pages[2].Width.Point, precision: 2);
        using var answer = PdfReader.Open(result.AnswerPdfPath!, PdfDocumentOpenMode.Import);
        Assert.Single(answer.Pages);
    }

    [Fact]
    public void Generate_RecordsMissingSourceAndContinuesWithSuccessfulJobs()
    {
        var source = CreatePdf("source.pdf", 1);
        var result = PdfMaterialGenerationService.Generate(
            [new("A", "教材", Path.Combine(_folder, "missing.pdf"), "", false),
             new("B", "教材", source, "", false)], _folder, duplex: false);
        Assert.Single(result.Outputs);
        var error = Assert.Single(result.Errors);
        Assert.Equal("A", error.StudentName);
        Assert.Equal(1, error.TestNumber);
        Assert.Contains("missing.pdf", error.TechnicalDetails);
        Assert.NotNull(result.ProblemPdfPath);
    }

    [Fact]
    public void Generate_UsesNewFolderForEveryRun()
    {
        var source = CreatePdf("source.pdf", 1);
        PdfMaterialRequest[] requests = [new("A", "段階突破", source, "", false)];
        var first = PdfMaterialGenerationService.Generate(requests, _folder, duplex: false);
        var second = PdfMaterialGenerationService.Generate(requests, _folder, duplex: false);
        Assert.NotEqual(first.OutputFolder, second.OutputFolder);
        Assert.True(File.Exists(first.ProblemPdfPath));
    }

    [Fact]
    public void MergeFiles_RejectsExistingDestination()
    {
        var source = CreatePdf("source.pdf", 1);
        var original = File.ReadAllBytes(source);
        Assert.Throws<IOException>(() => PdfMaterialGenerationService.MergeFiles([source], source, duplex: false));
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    private string CreatePdf(string name, int pageCount)
    {
        Directory.CreateDirectory(_folder);
        var path = Path.Combine(_folder, name);
        using var pdf = new PdfDocument();
        for (var index = 1; index <= pageCount; index++)
        {
            var page = pdf.AddPage();
            page.Width = XUnit.FromPoint(300 + index);
            page.Height = XUnit.FromPoint(500);
        }
        pdf.Save(path);
        return path;
    }

    public void Dispose()
    {
        if (Directory.Exists(_folder))
        {
            Directory.Delete(_folder, recursive: true);
        }
    }
}
