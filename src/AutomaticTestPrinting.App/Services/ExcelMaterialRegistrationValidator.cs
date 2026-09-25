using System.Runtime.InteropServices;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.App.Services;

public sealed record ExcelMaterialValidationResult(bool IsValid, string Message);

public static class ExcelMaterialRegistrationValidator
{
    private const int AutomationSecurityForceDisable = 3;

    public static Task<ExcelMaterialValidationResult> ValidateAsync(
        string workbookPath,
        ExcelTemplateProfile referenceProfile)
    {
        var completion = new TaskCompletionSource<ExcelMaterialValidationResult>(
            TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                completion.TrySetResult(Validate(workbookPath, referenceProfile));
            }
            catch (Exception exception)
            {
                completion.TrySetResult(new ExcelMaterialValidationResult(
                    false,
                    $"Excelファイルを確認できませんでした：{exception.Message}"));
            }
        })
        {
            IsBackground = true,
            Name = "Excel material validation"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private static ExcelMaterialValidationResult Validate(
        string workbookPath,
        ExcelTemplateProfile profile)
    {
        object? excel = null;
        object? workbooks = null;
        object? workbook = null;
        object? worksheets = null;
        var openedSheets = new List<object>();
        object? startRange = null;
        object? endRange = null;

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
            excelApplication.EnableEvents = false;
            excelApplication.AutomationSecurity = AutomationSecurityForceDisable;

            workbooks = excelApplication.Workbooks;
            workbook = ((dynamic)workbooks).Open(
                workbookPath,
                UpdateLinks: 0,
                ReadOnly: true,
                IgnoreReadOnlyRecommended: true,
                AddToMru: false);
            worksheets = ((dynamic)workbook).Worksheets;

            var requiredSheetNames = new[]
            {
                profile.WorkingSheetName,
                profile.QuestionListSheetName,
                profile.TeacherSheetName,
                profile.StudentSheetName
            }.Concat(profile.RangeUsesSectionMapping
                ? [profile.SectionMappingSheetName!]
                : [])
                .Distinct(StringComparer.OrdinalIgnoreCase);

            foreach (var sheetName in requiredSheetNames)
            {
                var sheet = ((dynamic)worksheets)[sheetName];
                openedSheets.Add(sheet);
            }

            if (!profile.RangeUsesSectionMapping)
            {
                dynamic workingSheet = openedSheets[0];
                startRange = workingSheet.Range[profile.RangeStartCell];
                endRange = workingSheet.Range[profile.RangeEndCell];
            }

            return new ExcelMaterialValidationResult(
                true,
                $"「{profile.DisplayName}」と同じ形式で利用できます。");
        }
        catch (COMException)
        {
            return new ExcelMaterialValidationResult(
                false,
                $"「{profile.DisplayName}」と同じシート構成ではありません。近い形式を選び直してください。");
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

            ReleaseComObject(endRange);
            ReleaseComObject(startRange);
            foreach (var sheet in openedSheets.AsEnumerable().Reverse())
            {
                ReleaseComObject(sheet);
            }
            ReleaseComObject(worksheets);
            ReleaseComObject(workbook);
            ReleaseComObject(workbooks);
            ReleaseComObject(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }
}
