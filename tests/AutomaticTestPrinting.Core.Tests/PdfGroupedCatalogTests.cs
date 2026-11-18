using System.Text.Json;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class PdfGroupedCatalogTests
{
    [Fact]
    public void RepeatedNumbersRequireExplicitPartAndKeepIndependentMaps()
    {
        var entry=Assert.Single(SharedMaterialCatalog.Load().Items,e=>e.Id=="shared-pdf-313");
        Assert.Equal(2,entry.PdfVariants.Count);
        Assert.Equal("1-1",entry.PdfVariants[0].Questions[0].Pages);
        Assert.Equal("14-14",entry.PdfVariants[1].Questions[0].Pages);
        var ready=PdfMaterialCatalog.Prepare(new(entry.DisplayName,"1-1",1,1,true),[]);
        Assert.True(ready.IsSupported);
        Assert.False(ready.IsValid);
        Assert.Contains("区別して指定",ready.Message);
        Assert.Contains("古文",ready.Message);
        Assert.Contains("漢文",ready.Message);
        Assert.DoesNotContain(SharedMaterialCatalog.CreatePdfProfiles(),p=>p.Id==entry.Id);
    }

    [Fact]
    public void NewMappingsKeepVerifiedPhysicalBoundsAndQuestionOnlyOutput()
    {
        var catalog=SharedMaterialCatalog.Load();
        Assert.Equal(58,SharedMaterialCatalog.CreatePdfProfiles().Count(p=>!p.IsStageTest));
        foreach(var e in catalog.Items.Where(e=>e.Kind=="PDF"&&e.AutoEnabled))
        foreach(var q in e.Questions.Concat(e.PdfVariants.SelectMany(v=>v.Questions)))
        {
            Assert.Contains(q.FileName,e.FileNames);
            Assert.NotNull(q.ExpectedPageCount);
            Assert.NotEmpty(PdfMaterialGenerationService.ParsePages(q.Pages,q.ExpectedPageCount!.Value));
            Assert.Empty(q.AnswerFileName);
        }
        var rules=Assert.Single(catalog.Items,e=>e.DisplayName=="(構文解釈)関正生のThe Rules 英語長文問題集3");
        Assert.Equal("28-35",rules.Questions.Single(q=>q.Number==5).Pages);
        var world=Assert.Single(catalog.Items,e=>e.DisplayName=="世界史用語マルチ・トレーニング STEP2");
        Assert.Equal("57-58",world.Questions.Single(q=>q.Number==57).Pages);
        Assert.Equal("59-59",world.Questions.Single(q=>q.Number==58).Pages);
    }

    [Fact]
    public void YearMappingsRetainRealYearsAndFacultyNames()
    {
        var entry=Assert.Single(SharedMaterialCatalog.Load().Items,e=>e.PdfVariants.Any(v=>v.DisplayName.Contains("阪大")));
        Assert.Equal(5,entry.PdfVariants.Count);
        Assert.All(entry.PdfVariants,v=>{
            Assert.True(v.AllowSparseRanges);
            Assert.Equal(Enumerable.Range(2008,15),v.Questions.Select(q=>q.Number));
            Assert.Contains("年度指定",v.DisplayName);
        });
    }

    [Fact]
    public void GroupedCatalogRejectsAmbiguousOrInvalidSourceDefinitions()
    {
        var profile=new PdfMaterialProfile("part","教材（部1）",[],[new(1,"book.pdf","1",ExpectedPageCount:2)]);
        var entry=new SharedMaterialEntry{Id="root",Kind="PDF",DisplayName="教材",FileNames=["book.pdf"],AutoEnabled=true,PdfVariants=[profile]};
        var db=new SharedMaterialDatabase{SchemaVersion=1,Items=[entry]};
        Assert.Single(SharedMaterialCatalog.Parse(JsonSerializer.Serialize(db)).Items);
        foreach(var invalid in new[]{
            entry with{Questions=[new(1,"book.pdf","1")]},
            entry with{PdfVariants=[profile with{IsStageTest=true}]},
            entry with{PdfVariants=[profile with{Questions=[new(1,"other.pdf","1",ExpectedPageCount:2)]}]},
            entry with{PdfVariants=[profile with{Questions=[]}]}
        }) Assert.Throws<InvalidDataException>(()=>SharedMaterialCatalog.Parse(JsonSerializer.Serialize(db with{Items=[invalid]})));
        Assert.Throws<ArgumentException>(()=>SharedMaterialCatalog.Parse(JsonSerializer.Serialize(db with{
            Items=[entry with{PdfVariants=[profile with{Questions=[new(1,"book.pdf","3",ExpectedPageCount:2)]}]}]})));
    }
}
