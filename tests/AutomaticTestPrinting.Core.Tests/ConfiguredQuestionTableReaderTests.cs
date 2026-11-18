using System.IO.Compression;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.Core.Tests;

public sealed class ConfiguredQuestionTableReaderTests
{
    [Fact]
    public void KeywordCatalogNameResolvesToTheMeaningsLayout()
    {
        var profiles = ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory, "materials.json")).Profiles;
        var request = new NormalTestRequestCandidate("現代文キーワード読解（キーワード）", "73-87", 10, 1, true);
        var result = ExcelTemplateCatalog.Prepare(request, Array.Empty<string>(), profiles: profiles);
        Assert.True(result.IsSupported);
        Assert.Equal("modern-japanese-keyword-meanings", result.Profile?.Id);
        Assert.DoesNotContain("作成設定は未登録", result.Message);
    }

    [Fact]
    public void JapaneseHistoryFiltersSerialNumbersWhileKeepingBookLabels()
    {
        var profiles = ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory, "materials.json")).Profiles;
        var profile = Assert.Single(profiles, p => p.Id == "individual-japanese-history-third-edition");
        Assert.Equal(3493, profile.MaximumQuestionNumber);
        Assert.Equal(100, profile.EffectiveMaximumQuestionCount);
        Assert.Equal(10, profile.OriginalGrammarQuestionsPerPage);
        object[,] values = {
            { "乱数", "番号", "難易度", "問題", "解答", "" },
            { 0d, "49-9", "A", "Q1", "A1", 100d },
            { 0d, "49-9①", "B", "Q2", "A2", 101d },
            { 0d, "総-1", "C", "Q3", "A3", 102d },
            { 0d, "", "", "", "", 103d }
        };
        var questions = ConfiguredQuestionTableReader.Read(values, profile, 101, 103);
        Assert.Equal([101, 102], questions.Select(q => q.SourceNumber));
        Assert.Equal(["49-9①", "総-1"], questions.Select(q => q.Number));
        Assert.Equal(["Q2", "Q3"], questions.Select(q => q.Question));
        Assert.Equal(["A2", "A3"], questions.Select(q => q.Answer));
    }

    [Fact]
    public void UnitMappingConvertsUnitsToWordNumbersAndRejectsBrokenRanges()
    {
        object[,] mapping={{1d,8d},{9d,15d},{16d,19d}};
        Assert.Equal((9,19),UnitQuestionRange.Resolve(mapping,2,3));
        mapping[2,0]=9d;
        Assert.Throws<InvalidOperationException>(()=>UnitQuestionRange.Resolve(mapping,2,3));
    }

    [Fact]
    public void ThemePrefixFiltersWholeThemeInsteadOfTheFirstFewSerialNumbers()
    {
        object[,] values={{"乱数","番号","テーマ番号","問題","解答"},{0d,10d,"1-1","Q1","A1"},{0d,11d,"1-2","Q2","A2"},{0d,12d,"2-1","Q3","A3"}};
        var profile=new ExcelTemplateProfile{SourceQuestionColumn="D",SourceAnswerColumn="E",SourceRangeColumn="C",SourceRangeUsesPrefix=true};
        var theme=ConfiguredQuestionTableReader.Read(values,profile,1,1);
        Assert.Equal([10,11],theme.Select(q=>q.SourceNumber));
        Assert.Equal("Q3",Assert.Single(ConfiguredQuestionTableReader.Read(values,profile,2,2)).Question);
    }

    [Fact]
    public void HistoryTargetUsesChapterReferenceAsDisplayNumberRatherThanQuestion()
    {
        var profiles = ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory, "materials.json")).Profiles;
        var profile = Assert.Single(profiles, p => p.Id == "batch-7b2c3a29342c86c2");
        Assert.Equal("C", profile.SourceDisplayNumberColumn);
        Assert.Equal("D", profile.SourceQuestionColumn);
        Assert.Equal("E", profile.SourceAnswerColumn);
        object[,] values = { { "乱数", "番号", "", "", "" }, { 0d, 1d, "1章-1-1", "最古の人類を（　）という。", "猿人" } };
        var question = Assert.Single(ConfiguredQuestionTableReader.Read(values, profile, 1, 1));
        Assert.Equal("1章-1-1", question.Number);
        Assert.Equal("最古の人類を（　）という。", question.Question);
        Assert.Equal("猿人", question.Answer);
    }

    [Fact]
    public void UnitCandidateCountUsesRegisteredQuestionCount()
    {
        var profile=new ExcelTemplateProfile{RangeUsesSectionMapping=true,RangeQuestionCounts=new Dictionary<int,int>{{1,8},{2,7},{3,4}}};
        Assert.Equal(11,UnitQuestionRange.CandidateCount(profile,2,3));
    }

    [Fact]
    public void DerivativeWordsKeepParentRangeAndReceiveDistinctPrintedNumbers()
    {
        object[,] values={{"乱数","番号","単語","訳"},{0d,1d,"create","作る"},{0d,1d,"creation","創造"},{0d,1d,"creation","創造"},{0d,2d,"follow","従う"}};
        var profile=new ExcelTemplateProfile{GroupDuplicateQuestionNumbers=true};
        var questions=ConfiguredQuestionTableReader.Read(values,profile,1,2);
        Assert.Equal(["1(1)","1(2)","2"],questions.Select(q=>q.Number));
        Assert.Equal([1,1,2],questions.Select(q=>q.SourceNumber));
        Assert.Equal("follow",Assert.Single(ConfiguredQuestionTableReader.Read(values,profile,2,2)).Question);
    }

    [Theory]
    [InlineData("ターゲット1900派生語")]
    [InlineData("システム英単語派生語")]
    [InlineData("必携英単語LEAP派生語")]
    public void DerivativeAliasesDoNotSelectTheMainVocabularyWorkbook(string name)
    {
        var profiles=ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory,"materials.json")).Profiles;
        var result=Assert.Single(ExcelMaterialNameMatcher.FindBestMatches(profiles,name));
        Assert.True(result.GroupDuplicateQuestionNumbers);
        Assert.True(result.AdditionalQuestionCounts.Count>0);
    }

    [Fact]
    public void BatchProfilesKeepEstablishedVintageAndFinalPolarisPagination()
    {
        var profiles=ExcelTemplateProfileStore.LoadFromFile(Path.Combine(AppContext.BaseDirectory,"materials.json")).Profiles;
        var variants=profiles.Where(p=>p.UseStandardQuestionTable && (p.WorkbookFileName.Contains("Vintage",StringComparison.OrdinalIgnoreCase) || p.QuestionExpansionProfileId.Length>0)).ToArray();
        Assert.Equal(4,variants.Length);
        Assert.All(variants,p=>Assert.Equal(25,p.OriginalGrammarQuestionsPerPage));
        var polaris=Assert.Single(variants,p=>p.QuestionExpansionProfileId.Length>0);
        Assert.Equal(20,PolarisSubquestionExpander.GetAvailableCount(polaris,285,286));
    }

    [Fact]
    public void MissingQuestionsAreExcludedBeforeExcelStarts()
    {
        var folder=Path.Combine(Path.GetTempPath(),"MissingQuestions-"+Guid.NewGuid());
        Directory.CreateDirectory(folder);
        try
        {
            var profile=new ExcelTemplateProfile{Id="gap",DisplayName="Gap",WorkbookNameKeyword="Gap",WorkbookFileName="Gap.xlsm",MaterialNameKeywords=["Gap"],
                MinimumQuestionNumber=1,MaximumQuestionNumber=10,MaximumQuestionCount=100,WorkingSheetName="work",RangeStartCell="G1",RangeEndCell="G2",
                QuestionListSheetName="list",TeacherSheetName="teacher",StudentSheetName="student",UnavailableQuestionNumbers=[3,4]};
            var result=ExcelTemplateCatalog.Prepare(new("Gap","1-10",9,1,true),[folder],profiles:[profile]);
            Assert.False(result.IsValid);Assert.Contains("1-8問",result.Message);
        }
        finally{Directory.Delete(folder);}
    }

    [Theory]
    [InlineData("A", "B", "C")]
    [InlineData("A", "C", "D")]
    [InlineData("B", "C", "D")]
    [InlineData("B", "D", "E")]
    [InlineData("C", "D", "E")]
    public void ReadsConfiguredColumnsWithoutMistakingRandomOrDifficultyForTheAnswer(string number, string question, string answer)
    {
        var profile = new ExcelTemplateProfile { SourceNumberColumn = number, SourceQuestionColumn = question, SourceAnswerColumn = answer };
        var values = new object[3,8];
        values[0,number[0]-'A']="番号";values[0,question[0]-'A']="問題";values[0,answer[0]-'A']="解答";
        values[1,number[0]-'A']=10d;values[1,question[0]-'A']="Question";values[1,answer[0]-'A']="Answer";
        values[2,number[0]-'A']=11d;
        var result=Assert.Single(ConfiguredQuestionTableReader.Read(values,profile,10,11));
        Assert.Equal("Question",result.Question);Assert.Equal("Answer",result.Answer);
        Assert.Equal(2,result.SourceRow);
        Assert.Empty(ConfiguredQuestionTableReader.Read(values,profile,11,11));
    }

    [Fact]
    public void AcceptsIdenticalDuplicatesButStopsOnConflictingAnswers()
    {
        object[,] values={{"乱数","番号","問題","解答"},{0d,1d,"Q","A"},{0d,1d,"Q","A"}};
        Assert.Single(ConfiguredQuestionTableReader.Read(values,new(),1,1));
        values[2,3]="Different";
        Assert.Throws<InvalidOperationException>(()=>ConfiguredQuestionTableReader.Read(values,new(),1,1));
    }

    [Fact]
    public void RejectsTheAnswerOnlyOrRandomColumnAsQuestionContent()
    {
        object[,] values={{"番号","番号","解答","RND"},{1d,1d,"A",0.3d}};
        Assert.Throws<InvalidOperationException>(()=>ConfiguredQuestionTableReader.Read(values,new(),1,1));
    }

    [Fact]
    public void RichStylesRemainTiedToCellsEvenWhenTheSameTextHasDifferentUnderlining()
    {
        var path=Path.Combine(Path.GetTempPath(),Guid.NewGuid()+".xlsm");
        try
        {
            using(var archive=ZipFile.Open(path,ZipArchiveMode.Create))
            {
                Write("xl/workbook.xml","<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"問題解答リスト\" r:id=\"rId1\"/></sheets></workbook>");
                Write("xl/_rels/workbook.xml.rels","<Relationships><Relationship Id=\"rId1\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                Write("xl/sharedStrings.xml","<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><r><rPr><u/></rPr><t>A</t></r><r><t>B</t></r><rPh><t>phonetic</t></rPh></si><si><r><t>A</t></r><r><rPr><vertAlign val=\"subscript\"/><b/></rPr><t>B</t></r></si></sst>");
                Write("xl/worksheets/sheet1.xml","<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"2\"><c r=\"C2\" t=\"s\"><v>0</v></c><c r=\"D2\" t=\"s\"><v>1</v></c></row></sheetData></worksheet>");
                void Write(string name,string text){using var writer=new StreamWriter(archive.CreateEntry(name).Open());writer.Write(text);}
            }
            var result=WorkbookTextStyleReader.Read(path,"問題解答リスト");
            var question=Assert.Single(result["C2"]);Assert.Equal(1,question.Start);Assert.True(question.Underline);
            var answer=Assert.Single(result["D2"]);Assert.Equal(2,answer.Start);Assert.True(answer.Bold);Assert.Equal("subscript",answer.VerticalPosition);
        }
        finally{File.Delete(path);}
    }
}
