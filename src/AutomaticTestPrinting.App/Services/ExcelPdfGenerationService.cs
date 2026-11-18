using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.App.Services;

public static class ExcelPdfGenerationService
{
    private const int CalculationManual = -4135;
    private const int LineStyleNone = -4142;
    private const int LineStyleContinuous = 1;
    private const int BorderWeightThin = 2;
    private const int FixedFormatPdf = 0;
    private const int QualityStandard = 0;
    private const int AutomationSecurityForceDisable = 3;
    private const int GrammarQuestionsPerPage = 25;
    private const double GrammarRowHeight = 32;
    private const double OriginalGrammarQuestionFontSize = 16;

    public static Task<ExcelPdfGenerationResult> GenerateAsync(
        ExcelPdfGenerationRequest request,
        CancellationToken cancellationToken = default)
    {
        var completion = new TaskCompletionSource<ExcelPdfGenerationResult>(
            TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                completion.TrySetResult(Generate(request, cancellationToken));
            }
            catch (OperationCanceledException exception)
            {
                completion.TrySetCanceled(exception.CancellationToken);
            }
            catch (Exception exception)
            {
                completion.TrySetException(exception);
            }
        })
        {
            IsBackground = true,
            Name = "Excel PDF generation"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private static ExcelPdfGenerationResult Generate(
        ExcelPdfGenerationRequest request,
        CancellationToken cancellationToken)
    {
        ValidateRequest(request);
        Directory.CreateDirectory(request.OutputFolder);

        var temporaryFolder = Path.Combine(
            Path.GetTempPath(),
            "AutomaticTestPrinting",
            Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(temporaryFolder);
        var temporaryWorkbookPath = Path.Combine(
            temporaryFolder,
            Path.GetFileName(request.WorkbookPath));
        File.Copy(request.WorkbookPath, temporaryWorkbookPath, false);

        object? excel = null;
        object? workbooks = null;
        object? workbook = null;
        object? worksheets = null;
        object? backgroundSheet = null;
        object? sectionMappingSheet = null;
        object? questionListSheet = null;
        object? teacherSheet = null;
        object? studentSheet = null;
        ExcelPdfGenerationResult? outputPaths = null;
        var completed = false;

        try
        {
            var questionEmphasis = request.Profile.UseKanjiWritingLayout
                ? WorkbookDiscoveryService.ReadBoldTextSpans(temporaryWorkbookPath)
                : null;
            var sourceTextStyles = request.Profile.UseStandardQuestionTable
                ? WorkbookTextStyleReader.Read(temporaryWorkbookPath, request.Profile.QuestionListSheetName) : null;
            var excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException(
                    "Microsoft Excelが見つかりません。Excel 2021がインストールされているか確認してください。");

            excel = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Microsoft Excelを起動できませんでした。");
            dynamic excelApplication = excel;
            excelApplication.Visible = false;
            excelApplication.DisplayAlerts = false;
            excelApplication.ScreenUpdating = false;
            excelApplication.EnableEvents = false;
            excelApplication.AutomationSecurity = AutomationSecurityForceDisable;

            workbooks = excelApplication.Workbooks;
            dynamic workbookCollection = workbooks;
            workbook = workbookCollection.Open(
                temporaryWorkbookPath,
                UpdateLinks: 0,
                ReadOnly: false,
                IgnoreReadOnlyRecommended: true,
                AddToMru: false);
            excelApplication.Calculation = CalculationManual;

            dynamic openedWorkbook = workbook;
            worksheets = openedWorkbook.Worksheets;
            dynamic worksheetCollection = worksheets;
            ValidateRequiredWorksheets(worksheetCollection, request);
            backgroundSheet = worksheetCollection[request.Profile.WorkingSheetName];
            if (request.Profile.RangeUsesSectionMapping)
            {
                sectionMappingSheet = worksheetCollection[
                    request.Profile.SectionMappingSheetName
                    ?? throw new InvalidOperationException("教材の対応表設定がありません。")];
            }
            questionListSheet = worksheetCollection[request.Profile.QuestionListSheetName];
            if (request.Profile.UseStandardQuestionTable)
            {
                // Use clean output sheets so saved hidden rows, sheet protection, and
                // print buttons in the original template cannot affect the new table.
                var suffix = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture)[..8];
                teacherSheet = worksheetCollection.Add();
                ((dynamic)teacherSheet).Name = "ATP_Answer_" + suffix;
                studentSheet = worksheetCollection.Add();
                ((dynamic)studentSheet).Name = "ATP_Question_" + suffix;
                ValidateSheetIdentity(teacherSheet, "ATP_Answer_" + suffix);
                ValidateSheetIdentity(studentSheet, "ATP_Question_" + suffix);
            }
            else
            {
                teacherSheet = worksheetCollection[request.Profile.TeacherSheetName];
                studentSheet = worksheetCollection[request.Profile.StudentSheetName];
                ValidateSheetIdentity(teacherSheet, request.Profile.TeacherSheetName);
                ValidateSheetIdentity(studentSheet, request.Profile.StudentSheetName);
            }

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbook(
                excelApplication,
                backgroundSheet,
                sectionMappingSheet,
                questionListSheet,
                teacherSheet,
                studentSheet,
                request,
                questionEmphasis, sourceTextStyles);

            outputPaths = TestPdfOutputPaths.FindAvailable(request.OutputFolder, request.StudentName,
                request.MaterialName, $"{request.StartNumber}-{request.EndNumber}", request.QuestionCount, includeAnswers: true);
            ExportSheetToPdf(teacherSheet, outputPaths.AnswerPdfPath!);
            cancellationToken.ThrowIfCancellationRequested();
            ExportSheetToPdf(studentSheet, outputPaths.ProblemPdfPath);

            EnsurePdfCreated(outputPaths.AnswerPdfPath!);
            EnsurePdfCreated(outputPaths.ProblemPdfPath);
            completed = true;
            return outputPaths;
        }
        catch (Exception exception)
        {
            exception.Data["SourceWorkbookPath"] = request.WorkbookPath;
            throw;
        }
        finally
        {
            if (workbook is not null)
            {
                try
                {
                    ((dynamic)workbook).Close(SaveChanges: false);
                }
                catch (COMException)
                {
                }
            }

            if (excel is not null)
            {
                try
                {
                    ((dynamic)excel).Quit();
                }
                catch (COMException)
                {
                }
            }

            ReleaseComObject(studentSheet);
            ReleaseComObject(teacherSheet);
            ReleaseComObject(questionListSheet);
            ReleaseComObject(sectionMappingSheet);
            ReleaseComObject(backgroundSheet);
            ReleaseComObject(worksheets);
            ReleaseComObject(workbook);
            ReleaseComObject(workbooks);
            ReleaseComObject(excel);
            CollectReleasedComObjects();

            if (!completed && outputPaths is not null)
            {
                TryDeleteFile(outputPaths.ProblemPdfPath);
                if (outputPaths.AnswerPdfPath is not null) TryDeleteFile(outputPaths.AnswerPdfPath);
            }

            try
            {
                Directory.Delete(temporaryFolder, true);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }

    private static void PrepareWorkbook(
        dynamic excelApplication,
        dynamic backgroundSheet,
        dynamic? sectionMappingSheet,
        dynamic questionListSheet,
        dynamic teacherSheet,
        dynamic studentSheet,
        ExcelPdfGenerationRequest request,
        IReadOnlyDictionary<string, IReadOnlyList<(int Start, int Length)>>? questionEmphasis,
        IReadOnlyDictionary<string, IReadOnlyList<WorkbookTextStyle>>? sourceTextStyles)
    {
        var useOriginalGrammarLayout = request.Profile.Id == "vintage-fourth-edition" || request.Profile.UseOriginalGrammarLayout ||
            request.Profile.UseLandscapeReadingLayout || request.Profile.UsePortraitReadingLayout ||
            (request.Profile.UseStandardQuestionTable && !VerticalJapaneseQuestionLayout.UsesLayout(request.Profile));
        foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
        {
            if (request.Profile.UseStandardQuestionTable)
            {
                dynamic shapes = sheet.Shapes;
                try
                {
                    while ((int)shapes.Count > 0)
                    {
                        dynamic shape = shapes.Item(1);
                        try { shape.Delete(); }
                        finally { ReleaseComObject(shape); }
                    }
                }
                finally { ReleaseComObject(shapes); }
            }
            dynamic outputRange = sheet.Range[StackedQuestionAnswerLayout.UsesLayout(request.Profile) ? "B1:H201" : "B1:H101"];
            dynamic borders = outputRange.Borders;
            try
            {
                if (request.Profile.UseStandardQuestionTable && !VerticalJapaneseQuestionLayout.UsesLayout(request.Profile))
                {
                    dynamic used = sheet.UsedRange;
                    try { used.UnMerge(); used.Clear(); }
                    finally { ReleaseComObject(used); }
                }
                outputRange.ClearContents();
                if (!useOriginalGrammarLayout)
                    borders.LineStyle = LineStyleNone;
            }
            finally
            {
                ReleaseComObject(borders);
                ReleaseComObject(outputRange);
            }
        }

        List<QuestionData> questions;
        if (request.Profile.UseStandardQuestionTable)
        {
            dynamic used = questionListSheet.UsedRange;
            dynamic rows = used.Rows;
            dynamic? source = null;
            try
            {
                var last = (int)used.Row + (int)rows.Count - 1;
                source = questionListSheet.Range[$"A1:H{last}"];
                var start = request.StartNumber;
                var end = request.EndNumber;
                if (request.Profile.SectionMappingUsesRowPairs)
                {
                    dynamic mapping = sectionMappingSheet!.Range[$"A1:B{request.Profile.MaximumQuestionNumber}"];
                    try { (start,end) = UnitQuestionRange.Resolve((object[,])mapping.Value2,start,end); }
                    finally { ReleaseComObject(mapping); }
                }
                var candidates = ConfiguredQuestionTableReader.Read((object[,])source.Value2, request.Profile,
                    start, end).SelectMany(question =>
                        PolarisSubquestionExpander.Expand(PolarisSubquestionExpander.ExpansionId(request.Profile),
                            new(question.SourceNumber, question.Number, question.Question, question.Answer))
                        .Select(child => new QuestionData(child.Number, child.Question, child.Answer,
                            child.Number == question.Number ? question.SourceRow : 0))).ToArray();
                if (candidates.Length < request.QuestionCount)
                    throw new InvalidOperationException($"範囲 {request.StartNumber}-{request.EndNumber} の有効な問題は{candidates.Length}問です。問題数を減らすか、元データの欠番・問題と解答の列を確認してください。");
                Random.Shared.Shuffle(candidates);
                questions = candidates.Take(request.QuestionCount).ToList();
            }
            finally { ReleaseComObject(source); ReleaseComObject(rows); ReleaseComObject(used); }
        }
        else if (request.Profile.RangeUsesSectionMapping || request.Profile.QuestionListUsesSectionNumbers)
        {
            if (sectionMappingSheet is null && !request.Profile.QuestionListUsesSectionNumbers)
            {
                throw new InvalidOperationException("教材の対応表シートが見つかりません。");
            }

            questions = ReadQuestionsFromSectionRange(
                sectionMappingSheet,
                questionListSheet,
                request.StartNumber,
                request.EndNumber,
                request.QuestionCount,
                request.Profile.SectionMappingUsesGridPairs,
                request.Profile.SourceHasSeparateDisplayNumber,
                request.Profile.QuestionListUsesSectionNumbers);
        }
        else if (request.Profile.UsesDirectGrammarQuestions || request.Profile.UsesDirectReadingQuestions || request.Profile.UseOriginalWordPairLayout || request.Profile.UseKanjiWritingLayout ||
                 request.Profile.UseKanjiReadingLayout || request.Profile.ReadNumberedVocabularyDirectly)
        {
            questions = ReadQuestions(backgroundSheet, questionListSheet, request.QuestionCount,
                request.StartNumber, request.EndNumber, request.Profile.SourceHasSeparateDisplayNumber,
                request.Profile.UseOriginalWordPairLayout, request.Profile.Id);
        }
        else
        {
            SetCellValue(backgroundSheet, request.Profile.RangeStartCell, request.StartNumber);
            SetCellValue(backgroundSheet, request.Profile.RangeEndCell, request.EndNumber);
            excelApplication.CalculateFullRebuild();
            questions = ReadQuestions(
                backgroundSheet,
                questionListSheet,
                request.QuestionCount);
        }
        if (request.Profile.UsePortraitReadingLayout)
            questions = questions.Select(question => question with
            {
                Question = ReadingTextFormatter.Normalize(question.Question, reduceInlinePadding: true),
                Answer = ReadingTextFormatter.Normalize(question.Answer, reduceInlinePadding: false)
            }).ToList();
        if (useOriginalGrammarLayout)
        {
            ApplyOriginalGrammarTemplate(teacherSheet, questions, request, includeAnswers: true, sourceTextStyles: sourceTextStyles);
            ApplyOriginalGrammarTemplate(studentSheet, questions, request, includeAnswers: false, sourceTextStyles: sourceTextStyles);
            var printedNumbers = new List<string>(questions.Count);
            ReadQuestionNumbers(teacherSheet, "B", questions.Count, printedNumbers);
            if (!printedNumbers.SequenceEqual(questions.Select(question => Convert.ToString(question.Number, CultureInfo.InvariantCulture))))
                throw new InvalidOperationException("印刷シートの問題番号が選択した問題と一致しません。");
            ValidateStudentAnswerCellsAreEmpty(studentSheet, questions.Count, 0);
            return;
        }

        if (VerticalJapaneseQuestionLayout.UsesLayout(request.Profile))
        {
            var verticalQuestions = questions.Select(q => new StackedQuestion(q.Number, q.Question, q.Answer)).ToArray();
            var title = request.Profile.HeaderTitle ?? request.MaterialName;
            var range = $"{request.StartNumber}-{request.EndNumber}";
            VerticalJapaneseExcelLayout.Apply(teacherSheet, verticalQuestions, title, range, includeAnswers: true,
                questionEmphasis: questionEmphasis, balancedKanjiRows: request.Profile.UseKanjiReadingLayout);
            VerticalJapaneseExcelLayout.Apply(studentSheet, verticalQuestions, title, range, includeAnswers: false,
                questionEmphasis: questionEmphasis, balancedKanjiRows: request.Profile.UseKanjiReadingLayout);
            return;
        }
        var grammar = request.Profile.UsesDirectGrammarQuestions;
        var headerSize = grammar ? 12 : 20;
        var questionsPerPage = request.Profile.UseOriginalWordPairLayout ? request.Profile.OriginalWordPairQuestionsPerPage : 50;
        var firstPageCount = grammar ? request.QuestionCount : Math.Min(request.QuestionCount, questionsPerPage);
        var secondPageCount = request.QuestionCount - firstPageCount;
        var titles = new object[,] { { "番号", "問題", "解答" } };

        if (request.Profile.UseOriginalWordPairLayout)
        {
            ExtendWordPairFormats(teacherSheet, firstPageCount, request.Profile.OriginalWordPairFormatRows);
            ExtendWordPairFormats(studentSheet, firstPageCount, request.Profile.OriginalWordPairFormatRows);
        }

        foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
        {
            SetRangeValue(sheet, "B1:D1", titles);
            SetRightHeader(
                sheet,
                $"&{headerSize} 範囲：{request.StartNumber}-{request.EndNumber}");
        }
        SetCenterHeader(teacherSheet, $"&{headerSize} {request.Profile.TeacherHeaderText}");
        SetCenterHeader(studentSheet, $"&{headerSize} {request.Profile.StudentHeaderText}");
        var headerTitle = request.Profile.HeaderTitle;
        if (grammar && string.IsNullOrWhiteSpace(headerTitle))
        {
            headerTitle = request.Profile.DisplayName;
        }
        if (!string.IsNullOrWhiteSpace(headerTitle))
        {
            var escapedTitle = headerTitle.Replace("&", "&&", StringComparison.Ordinal);
            SetLeftHeader(teacherSheet, $"&{headerSize} {escapedTitle}");
            SetLeftHeader(studentSheet, $"&{headerSize} {escapedTitle}");
        }

        if (request.Profile.SourceHasSeparateDisplayNumber)
        {
            foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
            {
                SetRangeNumberFormat(sheet, $"B2:B{firstPageCount + 1}", "@");
            }
        }

        SetRangeValue(
            teacherSheet,
            $"B2:D{firstPageCount + 1}",
            CreateOutputValues(questions, 0, firstPageCount, includeAnswers: true));
        SetRangeValue(
            studentSheet,
            $"B2:D{firstPageCount + 1}",
            CreateOutputValues(questions, 0, firstPageCount, includeAnswers: false));
        ApplyBorders(teacherSheet.Range[$"B1:D{firstPageCount + 1}"]);
        ApplyBorders(studentSheet.Range[$"B1:D{firstPageCount + 1}"]);

        if (secondPageCount > 0)
        {
            foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
            {
                SetRangeValue(sheet, "F1:H1", titles);
            }

            SetRangeValue(
                teacherSheet,
                $"F2:H{secondPageCount + 1}",
                CreateOutputValues(questions, firstPageCount, secondPageCount, includeAnswers: true));
            SetRangeValue(
                studentSheet,
                $"F2:H{secondPageCount + 1}",
                CreateOutputValues(questions, firstPageCount, secondPageCount, includeAnswers: false));
            ApplyBorders(teacherSheet.Range[$"F1:H{secondPageCount + 1}"]);
            ApplyBorders(studentSheet.Range[$"F1:H{secondPageCount + 1}"]);
        }

        if (request.Profile.UseOriginalWordPairLayout)
        {
            ApplyOriginalWordPairPrintSettings(teacherSheet, firstPageCount, secondPageCount, request, includeAnswers: true);
            ApplyOriginalWordPairPrintSettings(studentSheet, firstPageCount, secondPageCount, request, includeAnswers: false);
        }
        else
        {
            var printArea = secondPageCount > 0 ? "B1:H51" : $"B1:D{firstPageCount + 1}";
            SetPrintArea(teacherSheet, printArea);
            SetPrintArea(studentSheet, printArea);
        }

        if (request.Profile.UseWideAnswerLayout || grammar)
        {
            ApplyWideAnswerLayout(
                teacherSheet,
                studentSheet,
                firstPageCount,
                grammar ? 64 : request.Profile.QuestionColumnWidth,
                grammar ? 28 : request.Profile.AnswerColumnWidth,
                request.Profile.OutputRowHeight,
                request.Profile.AutoFitOutputRows,
                request.Profile.FitToSinglePageTall,
                grammar);
        }

        ValidateGeneratedQuestions(teacherSheet, firstPageCount, secondPageCount);
        ValidateStudentAnswerCellsAreEmpty(studentSheet, firstPageCount, secondPageCount);
    }

    private static void ValidateRequiredWorksheets(dynamic worksheets, ExcelPdfGenerationRequest request)
    {
        var actual = new List<string>();
        for (var i = 1; i <= (int)worksheets.Count; i++)
        {
            dynamic sheet = worksheets[i];
            try { actual.Add((string)sheet.Name); }
            finally { ReleaseComObject(sheet); }
        }
        var expected = new[] { request.Profile.WorkingSheetName, request.Profile.QuestionListSheetName,
            request.Profile.TeacherSheetName, request.Profile.StudentSheetName }
            .Concat(request.Profile.RangeUsesSectionMapping ? new[] { request.Profile.SectionMappingSheetName ?? "" } : Array.Empty<string>());
        var missing = expected.Where(n => !actual.Contains(n, StringComparer.OrdinalIgnoreCase)).Distinct().ToArray();
        if (missing.Length > 0)
            throw new InvalidOperationException($"使用Excelに必要なシートがありません。\n使用Excel：{request.WorkbookPath}\n不足シート：{string.Join(" / ", missing)}\n実際のシート：{string.Join(" / ", actual)}\n教材の版・登録形式を確認してください。");
    }


    private static void SetCellValue(dynamic sheet, string address, object value)
    {
        dynamic cell = sheet.Range[address];
        try
        {
            cell.Value2 = value;
        }
        finally
        {
            ReleaseComObject(cell);
        }
    }

    private static List<QuestionData> ReadQuestions(
        dynamic backgroundSheet,
        dynamic questionListSheet,
        int questionCount, int? directStart = null, int? directEnd = null, bool separateDisplayNumber = false,
        bool allowIdenticalRows = false, string? profileId = null)
    {
        dynamic? selectedNumberRange = directStart is null ? backgroundSheet.Range[$"A2:A{questionCount + 1}"] : null;
        dynamic usedRange = questionListSheet.UsedRange;
        dynamic usedRows = usedRange.Rows;
        var lastUsedRow = Convert.ToInt32(usedRange.Row, CultureInfo.InvariantCulture) +
            Convert.ToInt32(usedRows.Count, CultureInfo.InvariantCulture) - 1;
        var endColumn = separateDisplayNumber ? "E" : "D";
        dynamic sourceRange = questionListSheet.Range[$"B2:{endColumn}{lastUsedRow}"];
        try
        {
            object? selectedNumbers = selectedNumberRange?.Value2;
            var selectedNumberMatrix = selectedNumbers as object[,];
            var sourceValues = (object[,])sourceRange.Value2;
            var sourceByNumber = new Dictionary<int, QuestionData>();
            var questionColumn = separateDisplayNumber ? 3 : 2;
            for (var row = 1; row <= sourceValues.GetLength(0); row++)
            {
                if (!TryConvertQuestionNumber(sourceValues[row, 1], out var number))
                {
                    continue;
                }

                if (directStart.HasValue && (number < directStart.Value || number > directEnd.GetValueOrDefault()))
                    continue;
                var question = Convert.ToString(sourceValues[row, questionColumn], CultureInfo.InvariantCulture);
                var answer = Convert.ToString(sourceValues[row, questionColumn + 1], CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(question) || string.IsNullOrWhiteSpace(answer))
                {
                    continue;
                }

                var displayNumber = separateDisplayNumber
                    ? Convert.ToString(sourceValues[row, 2], CultureInfo.InvariantCulture)
                    : null;
                var data = new QuestionData(string.IsNullOrWhiteSpace(displayNumber) ? number : displayNumber, question, answer);
                if (!sourceByNumber.TryAdd(number, data) && !(allowIdenticalRows && sourceByNumber[number] == data))
                {
                    throw new InvalidOperationException(
                        $"Excelの問題解答リストに問題番号 {number} が重複しています。");
                }
            }

            if (directStart is not null)
            {
                var candidates = profileId is not null && PolarisSubquestionExpander.Supports(profileId)
                    ? sourceByNumber.SelectMany(pair => PolarisSubquestionExpander.Expand(profileId,
                        new PolarisSourceQuestion(pair.Key, Convert.ToString(pair.Value.Number, CultureInfo.InvariantCulture)!,
                            pair.Value.Question, pair.Value.Answer)))
                        .Select(child => new QuestionData(child.Number, child.Question, child.Answer)).ToArray()
                    : sourceByNumber.Values.ToArray();
                if (candidates.Length < questionCount)
                    throw new InvalidOperationException(
                        $"問題解答リストの {directStart}-{directEnd} に問題・解答が揃った行が {candidates.Length} 件しかありません。問題数を減らすか、使用Excelの版・形式を確認してください。");
                Random.Shared.Shuffle(candidates);
                return candidates.Take(questionCount).ToList();
            }

            var questions = new List<QuestionData>(questionCount);
            for (var row = 1; row <= questionCount; row++)
            {
                var selectedValue = selectedNumberMatrix is null
                    ? selectedNumbers
                    : selectedNumberMatrix[row, 1];
                if (!TryConvertQuestionNumber(selectedValue, out int number) ||
                    !sourceByNumber.TryGetValue(number, out var question))
                {
                    throw new InvalidOperationException(
                        $"Excelの問題解答リストから{row}問目を取得できませんでした。" +
                        $"選択値：{Convert.ToString(selectedValue, CultureInfo.InvariantCulture)}");
                }

                questions.Add(question);
            }

            return questions;
        }
        finally
        {
            ReleaseComObject(sourceRange);
            ReleaseComObject(usedRows);
            ReleaseComObject(usedRange);
            ReleaseComObject(selectedNumberRange);
        }
    }

    private static List<QuestionData> ReadQuestionsFromSectionRange(
        dynamic? sectionMappingSheet,
        dynamic questionListSheet,
        int startSection,
        int endSection,
        int questionCount,
        bool mappingUsesGridPairs,
        bool sourceHasSeparateDisplayNumber,
        bool numberColumnIsSection)
    {
        dynamic usedRange = questionListSheet.UsedRange;
        dynamic usedRows = usedRange.Rows;
        var lastUsedRow = Convert.ToInt32(usedRange.Row, CultureInfo.InvariantCulture) +
            Convert.ToInt32(usedRows.Count, CultureInfo.InvariantCulture) - 1;
        dynamic sourceRange = sourceHasSeparateDisplayNumber
            ? questionListSheet.Range[$"B2:E{lastUsedRow}"]
            : questionListSheet.Range[$"A2:C{lastUsedRow}"];
        try
        {
            SectionMappingData? gridMapping = mappingUsesGridPairs
                ? ReadGridSectionMapping((object)sectionMappingSheet!, startSection, endSection)
                : null;
            var (firstNumber, lastNumber) = numberColumnIsSection
                ? (startSection, endSection)
                : gridMapping is null
                    ? ReadRowSectionBounds((object)sectionMappingSheet!, startSection, endSection)
                    : (gridMapping.FirstNumber, gridMapping.LastNumber);

            var sourceValues = (object[,])sourceRange.Value2;
            var candidates = new List<QuestionData>();
            var seenNumbers = new HashSet<object>();
            for (var row = 1; row <= sourceValues.GetLength(0); row++)
            {
                if (!TryConvertQuestionNumber(sourceValues[row, 1], out var number) ||
                    number < firstNumber || number > lastNumber)
                {
                    continue;
                }

                object displayNumber = sourceHasSeparateDisplayNumber &&
                    gridMapping?.DisplayNumbers.TryGetValue(number, out var mappedNumber) == true
                        ? mappedNumber
                        : sourceHasSeparateDisplayNumber
                            ? Convert.ToString(sourceValues[row, 2], CultureInfo.InvariantCulture) ?? string.Empty
                            : number;
                var questionColumn = sourceHasSeparateDisplayNumber ? 3 : 2;
                var answerColumn = sourceHasSeparateDisplayNumber ? 4 : 3;
                var question = Convert.ToString(
                    sourceValues[row, questionColumn],
                    CultureInfo.InvariantCulture);
                var answer = Convert.ToString(
                    sourceValues[row, answerColumn],
                    CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(question) || string.IsNullOrWhiteSpace(answer))
                {
                    continue;
                }

                if (!seenNumbers.Add(numberColumnIsSection ? displayNumber : number))
                {
                    throw new InvalidOperationException(
                        $"Excelの問題リストに問題番号 {displayNumber} が重複しています。");
                }

                candidates.Add(new QuestionData(
                    string.IsNullOrWhiteSpace(Convert.ToString(displayNumber, CultureInfo.InvariantCulture))
                        ? number
                        : displayNumber,
                    question,
                    answer));
            }

            if (candidates.Count < questionCount)
            {
                throw new InvalidOperationException(
                    $"テーマ {startSection}-{endSection} からは{candidates.Count}問しか作成できません。問題数を減らしてください。");
            }

            for (var index = candidates.Count - 1; index > 0; index--)
            {
                var replacementIndex = Random.Shared.Next(index + 1);
                (candidates[index], candidates[replacementIndex]) =
                    (candidates[replacementIndex], candidates[index]);
            }

            return candidates.Take(questionCount).ToList();
        }
        finally
        {
            ReleaseComObject(sourceRange);
            ReleaseComObject(usedRows);
            ReleaseComObject(usedRange);
        }
    }

    private static (int FirstNumber, int LastNumber) ReadRowSectionBounds(
        object sectionMappingSheet,
        int startSection,
        int endSection)
    {
        dynamic sheet = sectionMappingSheet;
        dynamic firstNumberCell = sheet.Range[$"A{startSection}"];
        dynamic lastNumberCell = sheet.Range[$"B{endSection}"];
        try
        {
            object? firstNumberValue = firstNumberCell.Value2;
            object? lastNumberValue = lastNumberCell.Value2;
            if (!TryConvertQuestionNumber(firstNumberValue, out int firstNumber) ||
                !TryConvertQuestionNumber(lastNumberValue, out int lastNumber) ||
                firstNumber < 1 || lastNumber < firstNumber)
            {
                throw new InvalidOperationException(
                    $"テーマ {startSection}-{endSection} の対応範囲を取得できませんでした。");
            }

            return (firstNumber, lastNumber);
        }
        finally
        {
            ReleaseComObject(lastNumberCell);
            ReleaseComObject(firstNumberCell);
        }
    }

    private static SectionMappingData ReadGridSectionMapping(
        object sectionMappingSheet,
        int startSection,
        int endSection)
    {
        dynamic sheet = sectionMappingSheet;
        dynamic mappingRange = sheet.Range["B10:I42"];
        try
        {
            var values = (object[,])mappingRange.Value2;
            int? firstNumber = null;
            int? lastNumber = null;
            var displayNumbers = new Dictionary<int, string>();
            for (var row = 1; row <= values.GetLength(0); row++)
            {
                for (var sectionColumn = 1; sectionColumn <= 7; sectionColumn += 2)
                {
                    if (!TryConvertQuestionNumber(
                            values[row, sectionColumn],
                            out var sectionNumber) ||
                        !TryParseNumberRange(
                            values[row, sectionColumn + 1],
                            out var rangeStart,
                            out var rangeEnd))
                    {
                        continue;
                    }

                    if (sectionNumber == startSection)
                    {
                        firstNumber = rangeStart;
                    }

                    if (sectionNumber == endSection)
                    {
                        lastNumber = rangeEnd;
                    }

                    for (var number = rangeStart; number <= rangeEnd; number++)
                    {
                        displayNumbers[number] = $"{sectionNumber}-{number - rangeStart + 1}";
                    }
                }
            }

            if (firstNumber is null || lastNumber is null || lastNumber < firstNumber)
            {
                throw new InvalidOperationException(
                    $"見出し {startSection}-{endSection} の対応範囲を取得できませんでした。");
            }

            return new SectionMappingData(
                firstNumber.Value,
                lastNumber.Value,
                displayNumbers);
        }
        finally
        {
            ReleaseComObject(mappingRange);
        }
    }

    private static bool TryParseNumberRange(
        object? value,
        out int startNumber,
        out int endNumber)
    {
        startNumber = 0;
        endNumber = 0;
        var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        var pieces = text.Split('-', StringSplitOptions.TrimEntries);
        return pieces.Length == 2 &&
            int.TryParse(pieces[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out startNumber) &&
            int.TryParse(pieces[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out endNumber);
    }

    private static bool TryConvertQuestionNumber(object? value, out int number)
    {
        if (value is double numericValue)
        {
            number = Convert.ToInt32(numericValue, CultureInfo.InvariantCulture);
            return true;
        }

        return int.TryParse(
            Convert.ToString(value, CultureInfo.InvariantCulture),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out number);
    }

    private static object[,] CreateOutputValues(
        IReadOnlyList<QuestionData> questions,
        int startIndex,
        int count,
        bool includeAnswers)
    {
        var values = new object[count, 3];
        for (var index = 0; index < count; index++)
        {
            var question = questions[startIndex + index];
            values[index, 0] = question.Number;
            values[index, 1] = question.Question;
            values[index, 2] = includeAnswers ? question.Answer : string.Empty;
        }

        return values;
    }

    private static void SetRangeValue(dynamic sheet, string address, object value)
    {
        dynamic range = sheet.Range[address];
        try
        {
            range.Value2 = value;
        }
        finally
        {
            ReleaseComObject(range);
        }
    }

    private static void SetRightHeader(dynamic sheet, string value)
    {
        dynamic pageSetup = sheet.PageSetup;
        try
        {
            pageSetup.RightHeader = value;
        }
        finally
        {
            ReleaseComObject(pageSetup);
        }
    }

    private static void SetRangeNumberFormat(
        dynamic sheet,
        string address,
        string numberFormat)
    {
        dynamic range = sheet.Range[address];
        try
        {
            range.NumberFormat = numberFormat;
        }
        finally
        {
            ReleaseComObject(range);
        }
    }

    private static void SetCenterHeader(dynamic sheet, string value)
    {
        dynamic pageSetup = sheet.PageSetup;
        try
        {
            pageSetup.CenterHeader = value;
        }
        finally
        {
            ReleaseComObject(pageSetup);
        }
    }

    private static void SetLeftHeader(dynamic sheet, string value)
    {
        dynamic pageSetup = sheet.PageSetup;
        try
        {
            pageSetup.LeftHeader = value;
        }
        finally
        {
            ReleaseComObject(pageSetup);
        }
    }

    private static void SetPrintArea(dynamic sheet, string value)
    {
        dynamic pageSetup = sheet.PageSetup;
        try
        {
            pageSetup.PrintArea = value;
        }
        finally
        {
            ReleaseComObject(pageSetup);
        }
    }

    private static void ApplyWideAnswerLayout(
        dynamic teacherSheet,
        dynamic studentSheet,
        int rowCount,
        double questionColumnWidth,
        double answerColumnWidth,
        double outputRowHeight,
        bool autoFitOutputRows,
        bool fitToSinglePageTall,
        bool grammar)
    {
        foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
        {
            dynamic numberColumn = sheet.Columns["B"];
            dynamic questionColumn = sheet.Columns["C"];
            dynamic answerColumn = sheet.Columns["D"];
            dynamic layoutRange = sheet.Range["A1:Z101"];
            dynamic outputRange = sheet.Range[$"B1:D{rowCount + 1}"];
            dynamic font = outputRange.Font;
            dynamic pageSetup = sheet.PageSetup;
            try
            {
                layoutRange.UnMerge();
                outputRange.ClearFormats();
                numberColumn.ColumnWidth = 9;
                questionColumn.ColumnWidth = questionColumnWidth;
                answerColumn.ColumnWidth = answerColumnWidth;
                font.Name = "Yu Gothic UI";
                font.Size = 10;
                outputRange.WrapText = true;
                outputRange.Orientation = 0;
                outputRange.VerticalAlignment = -4160;
                sheet.ResetAllPageBreaks();
                pageSetup.Orientation = grammar ? 1 : 2;
                if (grammar)
                {
                    pageSetup.PaperSize = 9; // A4
                    pageSetup.TopMargin = 52;
                    pageSetup.BottomMargin = 28;
                    pageSetup.LeftMargin = 28;
                    pageSetup.RightMargin = 28;
                    pageSetup.HeaderMargin = 16;
                    pageSetup.PrintTitleRows = "$1:$1";
                    pageSetup.PrintTitleColumns = "";
                    // A fixed scale lets Excel honor the manual page breaks.
                    pageSetup.Zoom = 85;
                }
                else
                {
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesWide = 1;
                    pageSetup.FitToPagesTall = fitToSinglePageTall ? (object)1 : false;
                }
                pageSetup.CenterHorizontally = true;
            }
            finally
            {
                ReleaseComObject(pageSetup);
                ReleaseComObject(font);
                ReleaseComObject(outputRange);
                ReleaseComObject(layoutRange);
                ReleaseComObject(answerColumn);
                ReleaseComObject(questionColumn);
                ReleaseComObject(numberColumn);
            }

            dynamic headerRow = sheet.Rows[1];
            try
            {
                headerRow.RowHeight = 20;
            }
            finally
            {
                ReleaseComObject(headerRow);
            }

            if (grammar)
            {
                ApplyGrammarPagination(sheet, rowCount);
            }
            else if (autoFitOutputRows)
            {
                dynamic dataRange = sheet.Range[$"B2:D{rowCount + 1}"];
                dynamic dataRows = dataRange.Rows;
                try
                {
                    dataRows.AutoFit();
                }
                finally
                {
                    ReleaseComObject(dataRows);
                    ReleaseComObject(dataRange);
                }
            }
            else
            {
                for (var row = 2; row <= rowCount + 1; row++)
                {
                    dynamic outputRow = sheet.Rows[row];
                    try
                    {
                        outputRow.RowHeight = outputRowHeight;
                    }
                    finally
                    {
                        ReleaseComObject(outputRow);
                    }
                }
            }

            ApplyBorders(sheet.Range[$"B1:D{rowCount + 1}"]);
        }
    }

    private static void ApplyGrammarPagination(dynamic sheet, int questionCount)
    {
        FitGrammarColumn(sheet, "C", questionCount, GrammarRowHeight, 10);
        FitGrammarColumn(sheet, "D", questionCount, GrammarRowHeight, 10);
        SetGrammarPageBreaks(sheet, questionCount, GrammarRowHeight, GrammarQuestionsPerPage);
    }

    private static void ExtendWordPairFormats(dynamic sheet, int rowCount, int formatRows)
    {
        if (rowCount <= formatRows) return;
        dynamic original = sheet.Range[$"B2:H{formatRows + 1}"];
        try
        {
            for (var offset = formatRows; offset < rowCount; offset += formatRows)
            {
                dynamic destination = sheet.Range[$"B{offset + 2}"];
                try { original.Copy(Destination: destination); }
                finally { ReleaseComObject(destination); }
            }
        }
        finally
        {
            ReleaseComObject(original);
        }
    }

    private static void ApplyOriginalWordPairPrintSettings(dynamic sheet, int firstPageCount, int secondPageCount,
        ExcelPdfGenerationRequest request, bool includeAnswers)
    {
        var rowHeight = request.Profile.OriginalWordPairRowHeight;
        const double sourceFontSize = 16;
        FitGrammarColumn(sheet, "C", firstPageCount, rowHeight, sourceFontSize);
        if (includeAnswers)
            FitGrammarColumn(sheet, "D", firstPageCount, rowHeight, sourceFontSize);
        if (secondPageCount > 0)
        {
            FitGrammarColumn(sheet, "G", secondPageCount, rowHeight, sourceFontSize);
            if (includeAnswers)
                FitGrammarColumn(sheet, "H", secondPageCount, rowHeight, sourceFontSize);
        }

        var rowCount = Math.Max(firstPageCount, secondPageCount);
        dynamic rows = sheet.Rows[$"2:{rowCount + 1}"];
        dynamic setup = sheet.PageSetup;
        try
        {
            rows.RowHeight = rowHeight;
            sheet.ResetAllPageBreaks();
            setup.PaperSize = 9;
            setup.Orientation = 1;
            // Leave room for the heading row and rounding at the paper boundary.
            setup.Zoom = rowHeight > 31.5 && rowHeight < 40 ? 47 : 49;
            setup.PrintTitleRows = "";
            setup.PrintTitleColumns = "";
            setup.PrintArea = secondPageCount > 0
                ? $"B1:D{firstPageCount + 1},F1:H{secondPageCount + 1}"
                : $"B1:D{firstPageCount + 1}";
            setup.LeftHeader = "&20 " + (request.Profile.HeaderTitle ?? request.Profile.DisplayName).Replace("&", "&&", StringComparison.Ordinal);
            setup.CenterHeader = includeAnswers ? "&20 講師用" : "&20 生徒用";
            setup.RightHeader = $"&20 範囲：{request.StartNumber}-{request.EndNumber} ({request.QuestionCount}問)";
        }
        finally
        {
            ReleaseComObject(setup);
            ReleaseComObject(rows);
        }
    }

    private static void SetColumnWidth(dynamic sheet, string column, double width)
    {
        dynamic cells = sheet.Range[$"{column}:{column}"];
        try { cells.ColumnWidth = width; }
        finally { ReleaseComObject(cells); }
    }

    private static void ApplyOriginalGrammarTemplate(dynamic sheet, IReadOnlyList<QuestionData> questions,
        ExcelPdfGenerationRequest request, bool includeAnswers,
        IReadOnlyDictionary<string, IReadOnlyList<WorkbookTextStyle>>? sourceTextStyles = null)
    {
        // Reuse the original worksheet's first page for each subsequent page.
        var questionCount = questions.Count;
        var landscapeReading = request.Profile.UseLandscapeReadingLayout;
        var standard = request.Profile.UseStandardQuestionTable;
        var reading = landscapeReading || request.Profile.UsePortraitReadingLayout || standard;
        var rowsPerPage = landscapeReading ? 5 : request.Profile.OriginalGrammarQuestionsPerPage;
        var rowHeight = landscapeReading ? 182.85 : request.Profile.OriginalGrammarRowHeight;
        var readingQuestionFontSize = landscapeReading ? 24 : 16;
        var readingAnswerFontSize = standard ? 16 : landscapeReading ? 22 : 18;
        if (reading)
        {
            // Wide English text for students; more room for Japanese translations and
            // sentence analysis on the answer sheet. Both fit on a single A4 landscape page.
            SetColumnWidth(sheet, "B", 13);
            SetColumnWidth(sheet, "C", standard ? request.Profile.QuestionColumnWidth : landscapeReading ? (includeAnswers ? 112 : 155) : (includeAnswers ? 94 : 110));
            SetColumnWidth(sheet, "D", standard ? request.Profile.AnswerColumnWidth : landscapeReading ? (includeAnswers ? 158 : 115) : (includeAnswers ? 86 : 70));
            if (standard)
            {
                // ColumnWidth depends on the source workbook's Normal font. Measure
                // the resulting width so older workbooks do not spill onto a second sheet.
                dynamic tableColumns = sheet.Range["B:D"];
                try
                {
                    var actualWidth = Convert.ToDouble(tableColumns.Width, CultureInfo.InvariantCulture);
                    if (actualWidth > 1040)
                    {
                        var scale = 1040 / actualWidth;
                        SetColumnWidth(sheet, "B", 13 * scale);
                        SetColumnWidth(sheet, "C", request.Profile.QuestionColumnWidth * scale);
                        SetColumnWidth(sheet, "D", request.Profile.AnswerColumnWidth * scale);
                    }
                }
                finally { ReleaseComObject(tableColumns); }
                dynamic visibleRows = sheet.Rows[$"1:{questionCount + 1}"];
                dynamic visibleColumns = sheet.Columns["B:D"];
                try { visibleRows.Hidden = false; visibleColumns.Hidden = false; }
                finally { ReleaseComObject(visibleRows); ReleaseComObject(visibleColumns); }
            }
        }
        dynamic original = sheet.Range[$"B2:D{rowsPerPage + 1}"];
        try
        {
            for (var start = rowsPerPage; start < questionCount; start += rowsPerPage)
            {
                dynamic destination = sheet.Range[$"B{start + 2}"];
                try { original.Copy(Destination: destination); }
                finally { ReleaseComObject(destination); }
            }
        }
        finally
        {
            ReleaseComObject(original);
        }

        SetRangeValue(sheet, "B1:D1", new object[,] { { "番号", "問題", "解答" } });
        if (request.Profile.SourceHasSeparateDisplayNumber || standard)
            SetRangeNumberFormat(sheet, $"B2:B{questionCount + 1}", "@");
        SetRangeValue(sheet, $"B2:D{questionCount + 1}", CreateOutputValues(questions, 0, questionCount, includeAnswers));
        dynamic numberCells = sheet.Range[$"B2:B{questionCount + 1}"];
        dynamic numberFont = numberCells.Font;
        try
        {
            numberCells.WrapText = false;
            numberCells.ShrinkToFit = true;
            numberFont.Color = 0;
            numberFont.Size = 16;
            if (standard)
            {
                numberCells.Orientation = 0;
                numberCells.HorizontalAlignment = -4108;
                numberCells.VerticalAlignment = -4108;
            }
        }
        finally
        {
            ReleaseComObject(numberFont);
            ReleaseComObject(numberCells);
        }
        dynamic questionCells = sheet.Range[$"C2:C{questionCount + 1}"];
        dynamic questionFont = questionCells.Font;
        try
        {
            questionFont.Name = "Yu Gothic UI";
            questionFont.Size = reading ? readingQuestionFontSize : OriginalGrammarQuestionFontSize;
            questionCells.WrapText = true;
            if (reading) questionCells.ShrinkToFit = false;
            if (standard)
            {
                questionFont.Color = 0;
                questionFont.Bold = false;
                questionCells.Orientation = 0;
                questionCells.HorizontalAlignment = request.Profile.RightAlignTableText ? -4152 : -4131;
            }
            questionCells.VerticalAlignment = -4108; // xlCenter
        }
        finally
        {
            ReleaseComObject(questionFont);
            ReleaseComObject(questionCells);
        }
        if (reading && includeAnswers)
        {
            dynamic answerCells = sheet.Range[$"D2:D{questionCount + 1}"];
            dynamic answerFont = answerCells.Font;
            try
            {
                answerFont.Name = "Yu Gothic UI";
                answerFont.Size = readingAnswerFontSize;
                answerFont.Color = 0;
                answerFont.Underline = -4142;
                answerCells.WrapText = true;
                answerCells.ShrinkToFit = false;
                answerCells.VerticalAlignment = -4108;
                if (standard)
                {
                    answerCells.Orientation = 0;
                    answerCells.HorizontalAlignment = request.Profile.RightAlignTableText ? -4152 : -4131;
                }
            }
            finally
            {
                ReleaseComObject(answerFont);
                ReleaseComObject(answerCells);
            }
        }
        if (reading)
        {
            var outputStyles = new Dictionary<string, IReadOnlyList<WorkbookTextStyle>>(StringComparer.Ordinal);
            if (sourceTextStyles is not null)
                for (var index = 0; index < questions.Count; index++)
                {
                    var sourceRow = questions[index].SourceRow.ToString(CultureInfo.InvariantCulture);
                    if (sourceTextStyles.TryGetValue(request.Profile.SourceQuestionColumn + sourceRow, out var questionStyles))
                        outputStyles[$"C{index + 2}"] = questionStyles;
                    if (sourceTextStyles.TryGetValue(request.Profile.SourceAnswerColumn + sourceRow, out var answerStyles))
                        outputStyles[$"D{index + 2}"] = answerStyles;
                }
            FitReadingColumn(sheet, "C", questionCount, rowHeight, readingQuestionFontSize, rowsPerPage, outputStyles);
            if (includeAnswers) FitReadingColumn(sheet, "D", questionCount, rowHeight, readingAnswerFontSize, rowsPerPage, outputStyles);
        }
        else
        {
            FitGrammarColumn(sheet, "C", questionCount, rowHeight, OriginalGrammarQuestionFontSize,
                compactLongParagraphs: request.Profile.Id.StartsWith("final-polaris-", StringComparison.Ordinal));
            if (includeAnswers)
                FitGrammarColumn(sheet, "D", questionCount, rowHeight, 18,
                    compactLongParagraphs: request.Profile.Id.StartsWith("final-polaris-", StringComparison.Ordinal));
        }
        ApplyBorders(sheet.Range[$"B1:D{questionCount + 1}"]);

        sheet.ResetAllPageBreaks();
        dynamic page = sheet.PageSetup;
        try
        {
            page.PaperSize = 9; // A4
            page.Orientation = landscapeReading ? 2 : 1;
            page.Zoom = 50;
            page.PrintTitleRows = "$1:$1";
            page.PrintTitleColumns = "";
            page.PrintArea = $"B1:D{questionCount + 1}";
            if (standard)
            {
                page.TopMargin = request.Profile.StandardTableTopMargin;
                page.BottomMargin = 20;
                page.LeftMargin = 28;
                page.RightMargin = 28;
                page.HeaderMargin = 12;
                page.FooterMargin = 10;
                page.CenterHorizontally = true;
                page.CenterVertically = false;
                page.LeftFooter = "";
                page.CenterFooter = "&P / &N";
                page.RightFooter = "";
                dynamic header = sheet.Range["B1:D1"];
                dynamic font = header.Font;
                try { header.RowHeight = 24; header.Orientation = 0; font.Name = "Yu Gothic UI"; font.Size = 16; font.Bold = true; }
                finally { ReleaseComObject(font); ReleaseComObject(header); }
            }
            if (request.Profile.Id == "vintage-fourth-edition")
            {
                page.LeftHeader = "&24[4th] Vintage";
            }
            else if (!string.IsNullOrWhiteSpace(request.Profile.HeaderTitle))
            {
                page.LeftHeader = "&16 " + request.Profile.HeaderTitle.Replace("&", "&&", StringComparison.Ordinal);
            }
            var centerFormat = System.Text.RegularExpressions.Regex.Match((string)page.CenterHeader, @"^&\d{1,2}").Value;
            var rightFormat = System.Text.RegularExpressions.Regex.Match((string)page.RightHeader, @"^&\d{1,2}").Value;
            page.CenterHeader = $"{(centerFormat.Length > 0 ? centerFormat : "&20")}{(includeAnswers ? "講師用" : "生徒用")}";
            var unit = request.Profile.SectionMappingUsesRowPairs ? " 単元" : request.Profile.SourceRangeColumn.Length>0 ? " テーマ" : request.Profile.QuestionListUsesSectionNumbers ? " 講" : "";
            page.RightHeader = $"{(rightFormat.Length > 0 ? rightFormat : "&20")} 範囲：{request.StartNumber}-{request.EndNumber}{unit} ({questionCount}問)";
        }
        finally
        {
            ReleaseComObject(page);
        }
        SetGrammarPageBreaks(sheet, questionCount, rowHeight, rowsPerPage);
    }

    private static void SetGrammarPageBreaks(dynamic sheet, int questionCount, double rowHeight, int rowsPerPage)
    {
        dynamic rows = sheet.Rows[$"2:{questionCount + 1}"];
        try
        {
            rows.RowHeight = rowHeight;
            for (var row = rowsPerPage + 2; row <= questionCount + 1; row += rowsPerPage)
            {
                dynamic outputRow = sheet.Rows[row];
                try
                {
                    outputRow.PageBreak = -4135; // xlPageBreakManual
                }
                finally
                {
                    ReleaseComObject(outputRow);
                }
            }
        }
        finally
        {
            ReleaseComObject(rows);
        }
    }

    private static void ApplyTextStyles(dynamic cell, IReadOnlyList<WorkbookTextStyle> styles)
    {
        foreach (var style in styles)
        {
            dynamic characters = cell.Characters[style.Start, style.Length];
            dynamic font = characters.Font;
            try
            {
                if (style.Bold) font.Bold = true;
                if (style.Italic) font.Italic = true;
                if (style.Underline) font.Underline = 2;
                if (style.VerticalPosition == "superscript") font.Superscript = true;
                if (style.VerticalPosition == "subscript") font.Subscript = true;
            }
            finally { ReleaseComObject(font); ReleaseComObject(characters); }
        }
    }

    private static void FitReadingColumn(dynamic sheet, string column, int questionCount, double rowHeight, double defaultFontSize, int questionsPerPage,
        IReadOnlyDictionary<string, IReadOnlyList<WorkbookTextStyle>>? styles = null)
    {
        // AutoFit on an output row also measures its OTHER column. Measure in an empty
        // temporary sheet so a long answer cannot shrink an otherwise short question.
        dynamic workbook = sheet.Parent;
        dynamic worksheets = workbook.Worksheets;
        dynamic probe = worksheets.Add();
        dynamic measure = probe.Range["A1"];
        dynamic measureRows = measure.Rows;
        dynamic measureFont = measure.Font;
        dynamic outputColumn = sheet.Range[$"{column}:{column}"];
        try
        {
            measure.ColumnWidth = outputColumn.ColumnWidth;
            measure.WrapText = true;
            measure.ShrinkToFit = false;
            measureFont.Name = "Yu Gothic UI";
            for (var row = 2; row <= questionCount + 1; row++)
            {
                dynamic cell = sheet.Range[$"{column}{row}"];
                dynamic font = cell.Font;
                try
                {
                    measure.Value2 = cell.Value2;
                    measureFont.Bold = false;
                    measureFont.Italic = false;
                    measureFont.Underline = -4142;
                    measureFont.Superscript = false;
                    measureFont.Subscript = false;
                    IReadOnlyList<WorkbookTextStyle>? runs = null;
                    if (styles is not null && styles.TryGetValue($"{column}{row}", out runs)) ApplyTextStyles(measure, runs);
                    var low = 10; // Half-point units; minimum 5pt in the source sheet.
                    var high = (int)(defaultFontSize * 2);
                    measureFont.Size = defaultFontSize;
                    measureRows.AutoFit();
                    var best = (double)measure.RowHeight <= rowHeight ? high : 0;
                    if (best != 0) low = high + 1;
                    while (low <= high)
                    {
                        var middle = (low + high) / 2;
                        measureFont.Size = middle / 2.0;
                        measureRows.AutoFit();
                        if ((double)measure.RowHeight <= rowHeight)
                        {
                            best = middle;
                            low = middle + 1;
                        }
                        else high = middle - 1;
                    }
                    if (best == 0)
                        throw new InvalidOperationException($"{row - 1}問目の{(column == "C" ? "問題" : "解答")}が、1ページ{questionsPerPage}問の欄に収まりません。元Excelの文章量・改行を確認してください。");
                    font.Size = best / 2.0;
                    if (runs is not null) ApplyTextStyles(cell, runs);
                }
                finally
                {
                    ReleaseComObject(font);
                    ReleaseComObject(cell);
                }
            }
        }
        finally
        {
            ReleaseComObject(outputColumn);
            ReleaseComObject(measureFont);
            ReleaseComObject(measureRows);
            ReleaseComObject(measure);
            try { probe.Delete(); }
            finally
            {
                ReleaseComObject(probe);
                // Parent objects share the RCWs owned by Generate. Its final cleanup
                // releases them after export; releasing them here disconnects the workbook.
            }
        }
    }

    private static void FitGrammarColumn(dynamic sheet, string column, int questionCount, double rowHeight, double defaultFontSize,
        bool compactLongParagraphs = false)
    {
        dynamic cells = sheet.Range[$"{column}2:{column}{questionCount + 1}"];
        dynamic rows = cells.Rows;
        try
        {
            rows.AutoFit();
            if (rows.RowHeight is double uniformHeight && uniformHeight <= rowHeight)
                return;

            for (var row = 2; row <= questionCount + 1; row++)
            {
                dynamic cell = sheet.Range[$"{column}{row}"];
                try
                {
                    if ((double)cell.RowHeight <= rowHeight)
                        continue;

                    dynamic cellRow = cell.Rows;
                    dynamic font = cell.Font;
                    try
                    {
                        for (var size = defaultFontSize - 0.5; size >= 5 && (double)cell.RowHeight > rowHeight; size -= 0.5)
                        {
                            font.Size = size;
                            cellRow.AutoFit();
                        }
                        if ((double)cell.RowHeight > rowHeight)
                        {
                            var text = Convert.ToString(cell.Value2, CultureInfo.InvariantCulture) ?? "";
                            if (compactLongParagraphs && text.Contains('\n'))
                            {
                                cell.Value2 = text.Replace("\r", "", StringComparison.Ordinal).Replace("\n", "　", StringComparison.Ordinal);
                                for (var size = defaultFontSize; size >= 5 && (double)cell.RowHeight > rowHeight; size -= 0.5)
                                {
                                    font.Size = size;
                                    cellRow.AutoFit();
                                }
                            }
                        }
                        if ((double)cell.RowHeight > rowHeight)
                        {
                            cell.WrapText = false;
                            cell.ShrinkToFit = true;
                            font.Size = defaultFontSize;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(font);
                        ReleaseComObject(cellRow);
                    }
                }
                finally
                {
                    ReleaseComObject(cell);
                }
            }
        }
        finally
        {
            ReleaseComObject(rows);
            ReleaseComObject(cells);
        }
    }

    private static void ValidateSheetIdentity(dynamic sheet, string expectedName)
    {
        var actualName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
        if (!string.Equals(actualName, expectedName, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"Excelシートを正しく取得できませんでした。期待：{expectedName}、実際：{actualName}");
        }
    }

    private static void ValidateStudentAnswerCellsAreEmpty(
        dynamic studentSheet,
        int firstPageCount,
        int secondPageCount)
    {
        EnsureCellsAreEmpty(studentSheet, "D", firstPageCount);
        if (secondPageCount > 0)
        {
            EnsureCellsAreEmpty(studentSheet, "H", secondPageCount);
        }
    }

    private static void EnsureCellsAreEmpty(dynamic sheet, string column, int count)
    {
        for (var row = 2; row < count + 2; row++)
        {
            dynamic cell = sheet.Range[$"{column}{row}"];
            try
            {
                if (!string.IsNullOrEmpty(Convert.ToString(cell.Value2, CultureInfo.InvariantCulture)))
                {
                    throw new InvalidOperationException(
                        $"生徒用シートの解答欄に値が残っています：{column}{row}");
                }
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }
    }

    private static void ApplyBorders(dynamic range)
    {
        dynamic borders = range.Borders;
        try
        {
            borders.LineStyle = LineStyleContinuous;
            borders.Weight = BorderWeightThin;
        }
        finally
        {
            ReleaseComObject(borders);
            ReleaseComObject(range);
        }
    }

    private static void ValidateGeneratedQuestions(
        dynamic teacherSheet,
        int firstPageCount,
        int secondPageCount)
    {
        var numbers = new List<string>(firstPageCount + secondPageCount);
        ReadQuestionNumbers(teacherSheet, "B", firstPageCount, numbers);
        if (secondPageCount > 0)
        {
            ReadQuestionNumbers(teacherSheet, "F", secondPageCount, numbers);
        }

        if (numbers.Count != firstPageCount + secondPageCount)
        {
            throw new InvalidOperationException(
                "Excelで問題を正しく生成できませんでした。" +
                $"予定数：{firstPageCount + secondPageCount}、生成数：{numbers.Count}");
        }

        if (numbers.Distinct().Count() != numbers.Count)
        {
            var duplicateNumbers = numbers
                .GroupBy(number => number)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToArray();
            throw new InvalidOperationException(
                "Excelで問題を正しく生成できませんでした。" +
                $"生成数：{numbers.Count}、重複：{string.Join(",", duplicateNumbers)}");
        }
    }

    private static void ReadQuestionNumbers(
        dynamic sheet,
        string column,
        int count,
        ICollection<string> destination)
    {
        for (var row = 2; row < count + 2; row++)
        {
            dynamic cell = sheet.Range[$"{column}{row}"];
            try
            {
                var value = cell.Value2;
                var number = Convert.ToString(value, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(number))
                {
                    throw new InvalidOperationException(
                        "Excelの問題番号に計算エラーがあります。範囲を確認してください。");
                }

                destination.Add(number);
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }
    }

    private static void ExportSheetToPdf(dynamic sheet, string outputPath)
    {
        // The source workbook is saved with the teacher and student sheets grouped.
        // Replace the selection so Excel exports only the requested sheet.
        var temporaryPdf = Path.Combine(Path.GetTempPath(), $"AutomaticTestPrinting-{Guid.NewGuid():N}.pdf");
        try
        {
            sheet.Select(Replace: true);
            // Some Excel builds URL-escape spaces in the requested filename. Export with a
            // neutral name, then let the filesystem apply the human-readable output name.
            sheet.ExportAsFixedFormat(
                Type: FixedFormatPdf,
                Filename: temporaryPdf,
                Quality: QualityStandard,
                IncludeDocProperties: true,
                IgnorePrintAreas: false,
                OpenAfterPublish: false);
            EnsurePdfCreated(temporaryPdf);
            File.Move(temporaryPdf, outputPath);
        }
        finally
        {
            TryDeleteFile(temporaryPdf);
        }
    }

    private static void EnsurePdfCreated(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length == 0)
        {
            throw new InvalidOperationException(
                $"PDFを作成できませんでした：{Path.GetFileName(path)}");
        }
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static void ValidateRequest(ExcelPdfGenerationRequest request)
    {
        if (!File.Exists(request.WorkbookPath))
        {
            throw new FileNotFoundException("対応するExcelファイルが見つかりません。", request.WorkbookPath);
        }

        if (request.StartNumber < request.Profile.MinimumQuestionNumber ||
            request.EndNumber < request.StartNumber ||
            request.EndNumber > request.Profile.MaximumQuestionNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(request),
                "出題範囲が対応範囲外です。");
        }

        var availableCount = UnitQuestionRange.CandidateCount(request.Profile,request.StartNumber,request.EndNumber);
        if (request.QuestionCount < 1 ||
            request.QuestionCount > request.Profile.EffectiveMaximumQuestionCount ||
            request.QuestionCount > availableCount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(request),
                "問題数が対応範囲外です。");
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }

    private static void CollectReleasedComObjects()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private sealed record SectionMappingData(
        int FirstNumber,
        int LastNumber,
        IReadOnlyDictionary<int, string> DisplayNumbers);

    private sealed record QuestionData(object Number, string Question, string Answer, int SourceRow = 0);
}
