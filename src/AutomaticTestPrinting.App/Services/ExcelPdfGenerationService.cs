using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using AutomaticTestPrinting.Core.Models;

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

            outputPaths = CreateOutputPaths(request);
            ExportSheetToPdf(teacherSheet, outputPaths.AnswerPdfPath);
            cancellationToken.ThrowIfCancellationRequested();
            ExportSheetToPdf(studentSheet, outputPaths.ProblemPdfPath);

            EnsurePdfCreated(outputPaths.AnswerPdfPath);
            EnsurePdfCreated(outputPaths.ProblemPdfPath);
            completed = true;
            return outputPaths;
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
                TryDeleteFile(outputPaths.AnswerPdfPath);
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
        foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
        {
            dynamic outputRange = sheet.Range["B1:H101"];
            dynamic borders = outputRange.Borders;
            try
            {
                outputRange.ClearContents();
                borders.LineStyle = LineStyleNone;
            }
            finally
            {
                ReleaseComObject(borders);
                ReleaseComObject(outputRange);
            }
        }

        List<QuestionData> questions;
        if (request.Profile.RangeUsesSectionMapping)
        {
            if (sectionMappingSheet is null)
            {
                throw new InvalidOperationException("教材の対応表シートが見つかりません。");
            }

            questions = ReadQuestionsFromSectionRange(
                sectionMappingSheet,
                questionListSheet,
                request.StartNumber,
                request.EndNumber,
                request.QuestionCount);
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
        var firstPageCount = Math.Min(request.QuestionCount, 50);
        var secondPageCount = Math.Max(request.QuestionCount - 50, 0);
        var titles = new object[,] { { "番号", "問題", "解答" } };

        foreach (dynamic sheet in new[] { teacherSheet, studentSheet })
        {
            SetRangeValue(sheet, "B1:D1", titles);
            SetRightHeader(
                sheet,
                $"&20 範囲：{request.StartNumber}-{request.EndNumber}");
        }
        SetCenterHeader(teacherSheet, "&20 解答");
        SetCenterHeader(studentSheet, "&20 問題");

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

        if (request.Profile.UseWideAnswerLayout)
        {
            ApplyWideAnswerLayout(
                teacherSheet,
                studentSheet,
                firstPageCount,
                request.Profile.QuestionColumnWidth,
                request.Profile.AnswerColumnWidth,
                request.Profile.OutputRowHeight);
        }

        ValidateGeneratedQuestions(teacherSheet, firstPageCount, secondPageCount);
        ValidateStudentAnswerCellsAreEmpty(studentSheet, firstPageCount, secondPageCount);
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
        int questionCount)
    {
        dynamic selectedNumberRange = backgroundSheet.Range[$"A2:A{questionCount + 1}"];
        dynamic usedRange = questionListSheet.UsedRange;
        dynamic usedRows = usedRange.Rows;
        var lastUsedRow = Convert.ToInt32(usedRange.Row, CultureInfo.InvariantCulture) +
            Convert.ToInt32(usedRows.Count, CultureInfo.InvariantCulture) - 1;
        dynamic sourceRange = questionListSheet.Range[$"B2:D{lastUsedRow}"];
        try
        {
            object? selectedNumbers = selectedNumberRange.Value2;
            var selectedNumberMatrix = selectedNumbers as object[,];
            var sourceValues = (object[,])sourceRange.Value2;
            var sourceByNumber = new Dictionary<int, QuestionData>();
            for (var row = 1; row <= sourceValues.GetLength(0); row++)
            {
                if (!TryConvertQuestionNumber(sourceValues[row, 1], out var number))
                {
                    continue;
                }

                var question = Convert.ToString(sourceValues[row, 2], CultureInfo.InvariantCulture);
                var answer = Convert.ToString(sourceValues[row, 3], CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(question) || string.IsNullOrWhiteSpace(answer))
                {
                    continue;
                }

                if (!sourceByNumber.TryAdd(number, new QuestionData(number, question, answer)))
                {
                    throw new InvalidOperationException(
                        $"Excelの問題解答リストに問題番号 {number} が重複しています。");
                }
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
                        "Excelの問題・解答リストから出題内容を取得できませんでした。");
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
        dynamic sectionMappingSheet,
        dynamic questionListSheet,
        int startSection,
        int endSection,
        int questionCount)
    {
        dynamic firstNumberCell = sectionMappingSheet.Range[$"A{startSection}"];
        dynamic lastNumberCell = sectionMappingSheet.Range[$"B{endSection}"];
        dynamic usedRange = questionListSheet.UsedRange;
        dynamic usedRows = usedRange.Rows;
        var lastUsedRow = Convert.ToInt32(usedRange.Row, CultureInfo.InvariantCulture) +
            Convert.ToInt32(usedRows.Count, CultureInfo.InvariantCulture) - 1;
        dynamic sourceRange = questionListSheet.Range[$"A2:C{lastUsedRow}"];
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

            var sourceValues = (object[,])sourceRange.Value2;
            var candidates = new List<QuestionData>();
            var seenNumbers = new HashSet<int>();
            for (var row = 1; row <= sourceValues.GetLength(0); row++)
            {
                if (!TryConvertQuestionNumber(sourceValues[row, 1], out var number) ||
                    number < firstNumber || number > lastNumber)
                {
                    continue;
                }

                var question = Convert.ToString(sourceValues[row, 2], CultureInfo.InvariantCulture);
                var answer = Convert.ToString(sourceValues[row, 3], CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(question) || string.IsNullOrWhiteSpace(answer))
                {
                    continue;
                }

                if (!seenNumbers.Add(number))
                {
                    throw new InvalidOperationException(
                        $"Excelの単語リストに問題番号 {number} が重複しています。");
                }

                candidates.Add(new QuestionData(number, question, answer));
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
            ReleaseComObject(lastNumberCell);
            ReleaseComObject(firstNumberCell);
        }
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
        double outputRowHeight)
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
                pageSetup.Orientation = 2;
                pageSetup.Zoom = false;
                pageSetup.FitToPagesWide = 1;
                pageSetup.FitToPagesTall = false;
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

            for (var row = 1; row <= rowCount + 1; row++)
            {
                dynamic outputRow = sheet.Rows[row];
                try
                {
                    outputRow.RowHeight = row == 1 ? 20 : outputRowHeight;
                }
                finally
                {
                    ReleaseComObject(outputRow);
                }
            }

            ApplyBorders(sheet.Range[$"B1:D{rowCount + 1}"]);
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
        var numbers = new List<int>(firstPageCount + secondPageCount);
        ReadQuestionNumbers(teacherSheet, "B", firstPageCount, numbers);
        if (secondPageCount > 0)
        {
            ReadQuestionNumbers(teacherSheet, "F", secondPageCount, numbers);
        }

        if (numbers.Count != firstPageCount + secondPageCount ||
            numbers.Distinct().Count() != numbers.Count)
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
        ICollection<int> destination)
    {
        for (var row = 2; row < count + 2; row++)
        {
            dynamic cell = sheet.Range[$"{column}{row}"];
            try
            {
                var value = cell.Value2;
                if (value is not double numericValue)
                {
                    throw new InvalidOperationException(
                        "Excelの問題番号に計算エラーがあります。範囲を確認してください。");
                }

                destination.Add(Convert.ToInt32(numericValue, CultureInfo.InvariantCulture));
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }
    }

    private static ExcelPdfGenerationResult CreateOutputPaths(ExcelPdfGenerationRequest request)
    {
        var studentName = SanitizeFileName(request.StudentName, "氏名未確認");
        var materialName = SanitizeFileName(request.MaterialName, "教材");
        var baseName =
            $"{studentName}_{materialName}_{request.StartNumber}-{request.EndNumber}_{request.QuestionCount}問";
        if (baseName.Length > 120)
        {
            baseName = baseName[..120];
        }

        for (var suffix = 1; ; suffix++)
        {
            var suffixText = suffix == 1 ? string.Empty : $"_{suffix}";
            var problemPath = Path.Combine(
                request.OutputFolder,
                $"{baseName}{suffixText}_問題.pdf");
            var answerPath = Path.Combine(
                request.OutputFolder,
                $"{baseName}{suffixText}_解答.pdf");
            if (!File.Exists(problemPath) && !File.Exists(answerPath))
            {
                return new ExcelPdfGenerationResult(problemPath, answerPath);
            }
        }
    }

    private static string SanitizeFileName(string value, string fallback)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitized = new string(value
            .Trim()
            .Select(character => invalidCharacters.Contains(character) ? '_' : character)
            .ToArray());
        return string.IsNullOrWhiteSpace(sanitized) ? fallback : sanitized;
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

        var availableCount = request.Profile.RangeUsesSectionMapping
            ? request.Profile.MaximumQuestionCount
            : request.EndNumber - request.StartNumber + 1;
        if (request.QuestionCount < 1 ||
            request.QuestionCount > request.Profile.MaximumQuestionCount ||
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

    private sealed record QuestionData(int Number, string Question, string Answer);
}
