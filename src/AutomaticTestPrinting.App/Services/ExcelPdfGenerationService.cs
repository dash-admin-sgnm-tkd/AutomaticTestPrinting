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
            teacherSheet = worksheetCollection[request.Profile.TeacherSheetName];
            studentSheet = worksheetCollection[request.Profile.StudentSheetName];

            ValidateSheetIdentity(teacherSheet, request.Profile.TeacherSheetName);
            ValidateSheetIdentity(studentSheet, request.Profile.StudentSheetName);

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbook(
                excelApplication,
                backgroundSheet,
                sectionMappingSheet,
                questionListSheet,
                teacherSheet,
                studentSheet,
                request);

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
        ExcelPdfGenerationRequest request)
    {
        var useOriginalGrammarLayout = request.Profile.Id == "vintage-fourth-edition" || request.Profile.UseOriginalGrammarLayout;
        foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
        {
            dynamic outputRange = sheet.Range[StackedQuestionAnswerLayout.UsesLayout(request.Profile) ? "B1:H201" : "B1:H101"];
            dynamic borders = outputRange.Borders;
            try
            {
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
        if (request.Profile.RangeUsesSectionMapping || request.Profile.QuestionListUsesSectionNumbers)
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
        else if (request.Profile.UsesDirectGrammarQuestions)
        {
            questions = ReadQuestions(backgroundSheet, questionListSheet, request.QuestionCount,
                request.StartNumber, request.EndNumber, request.Profile.SourceHasSeparateDisplayNumber);
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
        if (useOriginalGrammarLayout)
        {
            ApplyOriginalGrammarTemplate(teacherSheet, questions, request, includeAnswers: true);
            ApplyOriginalGrammarTemplate(studentSheet, questions, request, includeAnswers: false);
            ValidateGeneratedQuestions(teacherSheet, questions.Count, 0);
            ValidateStudentAnswerCellsAreEmpty(studentSheet, questions.Count, 0);
            return;
        }

        if (VerticalJapaneseQuestionLayout.UsesLayout(request.Profile))
        {
            var verticalQuestions = questions.Select(q => new StackedQuestion(q.Number, q.Question, q.Answer)).ToArray();
            var title = request.Profile.HeaderTitle ?? request.MaterialName;
            var range = $"{request.StartNumber}-{request.EndNumber}";
            VerticalJapaneseExcelLayout.Apply(teacherSheet, verticalQuestions, title, range, includeAnswers: true);
            VerticalJapaneseExcelLayout.Apply(studentSheet, verticalQuestions, title, range, includeAnswers: false);
            return;
        }
        var grammar = request.Profile.UsesDirectGrammarQuestions;
        var headerSize = grammar ? 12 : 20;
        var firstPageCount = grammar ? request.QuestionCount : Math.Min(request.QuestionCount, 50);
        var secondPageCount = request.QuestionCount - firstPageCount;
        var titles = new object[,] { { "番号", "問題", "解答" } };

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
                CreateOutputValues(questions, 50, secondPageCount, includeAnswers: true));
            SetRangeValue(
                studentSheet,
                $"F2:H{secondPageCount + 1}",
                CreateOutputValues(questions, 50, secondPageCount, includeAnswers: false));
            ApplyBorders(teacherSheet.Range[$"F1:H{secondPageCount + 1}"]);
            ApplyBorders(studentSheet.Range[$"F1:H{secondPageCount + 1}"]);
        }

        var printArea = secondPageCount > 0
            ? "B1:H51"
            : $"B1:D{firstPageCount + 1}";
        SetPrintArea(teacherSheet, printArea);
        SetPrintArea(studentSheet, printArea);

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
        int questionCount, int? directStart = null, int? directEnd = null, bool separateDisplayNumber = false)
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
                if (!sourceByNumber.TryAdd(number, new QuestionData(
                    string.IsNullOrWhiteSpace(displayNumber) ? number : displayNumber, question, answer)))
                {
                    throw new InvalidOperationException(
                        $"Excelの問題解答リストに問題番号 {number} が重複しています。");
                }
            }

            if (directStart is not null)
            {
                var candidates = sourceByNumber.Values.ToArray();
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

    private static void ApplyOriginalGrammarTemplate(dynamic sheet, IReadOnlyList<QuestionData> questions,
        ExcelPdfGenerationRequest request, bool includeAnswers)
    {
        // Reuse the original worksheet's first page for each subsequent page.
        var questionCount = questions.Count;
        var rowsPerPage = request.Profile.OriginalGrammarQuestionsPerPage;
        var rowHeight = request.Profile.OriginalGrammarRowHeight;
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
        if (request.Profile.SourceHasSeparateDisplayNumber)
            SetRangeNumberFormat(sheet, $"B2:B{questionCount + 1}", "@");
        SetRangeValue(sheet, $"B2:D{questionCount + 1}", CreateOutputValues(questions, 0, questionCount, includeAnswers));
        dynamic questionCells = sheet.Range[$"C2:C{questionCount + 1}"];
        dynamic questionFont = questionCells.Font;
        try
        {
            questionFont.Name = "Yu Gothic UI";
            questionFont.Size = OriginalGrammarQuestionFontSize;
            questionCells.WrapText = true;
            questionCells.VerticalAlignment = -4108; // xlCenter
        }
        finally
        {
            ReleaseComObject(questionFont);
            ReleaseComObject(questionCells);
        }
        FitGrammarColumn(sheet, "C", questionCount, rowHeight, OriginalGrammarQuestionFontSize);
        if (includeAnswers)
            FitGrammarColumn(sheet, "D", questionCount, rowHeight, 18);
        ApplyBorders(sheet.Range[$"B1:D{questionCount + 1}"]);

        sheet.ResetAllPageBreaks();
        dynamic page = sheet.PageSetup;
        try
        {
            page.PaperSize = 9; // A4
            page.Orientation = 1;
            page.Zoom = 50;
            page.PrintTitleRows = "$1:$1";
            page.PrintTitleColumns = "";
            page.PrintArea = $"B1:D{questionCount + 1}";
            if (request.Profile.Id == "vintage-fourth-edition")
            {
                page.LeftHeader = "&24[4th] Vintage";
            }
            var centerFormat = System.Text.RegularExpressions.Regex.Match((string)page.CenterHeader, @"^&\d{1,2}").Value;
            var rightFormat = System.Text.RegularExpressions.Regex.Match((string)page.RightHeader, @"^&\d{1,2}").Value;
            page.CenterHeader = $"{(centerFormat.Length > 0 ? centerFormat : "&20")}{(includeAnswers ? "講師用" : "生徒用")}";
            var unit = request.Profile.QuestionListUsesSectionNumbers ? " 講" : "";
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

    private static void FitGrammarColumn(dynamic sheet, string column, int questionCount, double rowHeight, double defaultFontSize)
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
        sheet.Select(Replace: true);
        sheet.ExportAsFixedFormat(
            Type: FixedFormatPdf,
            Filename: outputPath,
            Quality: QualityStandard,
            IncludeDocProperties: true,
            IgnorePrintAreas: false,
            OpenAfterPublish: false);
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

        if (request.StartNumber < 1 ||
            request.EndNumber < request.StartNumber ||
            request.EndNumber > request.Profile.MaximumQuestionNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(request),
                "出題範囲が対応範囲外です。");
        }

        var availableCount = request.Profile.RangeUsesSectionMapping || request.Profile.QuestionListUsesSectionNumbers
            ? request.Profile.EffectiveMaximumQuestionCount
            : request.EndNumber - request.StartNumber + 1;
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

    private sealed record QuestionData(object Number, string Question, string Answer);
}
