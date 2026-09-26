using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.App.Models;

public sealed class RecognitionResultItem : INotifyPropertyChanged
{
    private string _studentName;
    private bool _isIncluded = true;

    private RecognitionResultItem(
        string sourcePath,
        string studentName,
        string normalTestSummary,
        string stageTestSummary,
        string orientationSummary,
        bool hasError,
        string errorType,
        string errorCode,
        string errorDetails,
        IEnumerable<EditableTestRequestItem> requests)
    {
        SourcePath = sourcePath;
        FileName = Path.GetFileName(sourcePath);
        _studentName = studentName;
        NormalTestSummary = normalTestSummary;
        StageTestSummary = stageTestSummary;
        OrientationSummary = orientationSummary;
        HasError = hasError;
        ErrorType = errorType;
        ErrorCode = errorCode;
        ErrorDetails = errorDetails;
        Requests = new ObservableCollection<EditableTestRequestItem>(requests);
        foreach (var request in Requests)
        {
            request.Edited += Request_Edited;
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? Edited;

    public string SourcePath { get; }
    public string FileName { get; }
    public string NormalTestSummary { get; }
    public string StageTestSummary { get; }
    public string OrientationSummary { get; }
    public bool HasError { get; }
    public string ErrorType { get; }
    public string ErrorCode { get; }
    public string ErrorDetails { get; }
    public ObservableCollection<EditableTestRequestItem> Requests { get; }

    public bool IsIncluded
    {
        get => _isIncluded;
        set
        {
            if (_isIncluded == value)
            {
                return;
            }

            _isIncluded = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsIncluded)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReady)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasIncludedRequests)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NeedsAttention)));
            Edited?.Invoke(this, EventArgs.Empty);
        }
    }

    public string StudentName
    {
        get => _studentName;
        set
        {
            if (string.Equals(_studentName, value, StringComparison.Ordinal))
            {
                return;
            }

            _studentName = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StudentName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReady)));
            Edited?.Invoke(this, EventArgs.Empty);
        }
    }

    public bool IsReady =>
        !IsIncluded ||
        (!HasError &&
         (!Requests.Any(request => request.IsIncluded) ||
          (!string.IsNullOrWhiteSpace(StudentName) &&
           Requests.All(request => request.IsReady))));

    public bool HasIncludedRequests =>
        IsIncluded && Requests.Any(request => request.IsIncluded);

    public bool NeedsAttention =>
        IsIncluded &&
        (HasError || Requests.Any(request => request.IsIncluded && !request.IsValid));

    public static RecognitionResultItem Success(RecognizedReport report, string? materialFolder)
    {
        var requests = report.TestRequests
            .Select((request, index) => new EditableTestRequestItem(request, index + 1, materialFolder))
            .ToArray();
        return new RecognitionResultItem(
            report.SourcePath,
            report.StudentName,
            report.NormalTestSummary,
            report.StageTestSummary,
            report.OrientationSummary,
            false,
            string.Empty,
            string.Empty,
            string.Empty,
            requests);
    }

    public static RecognitionResultItem Failure(string sourcePath, Exception exception) => new(
        sourcePath,
        "読み取りに失敗しました",
        exception.Message,
        "PDFを開けるか確認してください",
        string.Empty,
        true,
        exception.GetType().FullName ?? exception.GetType().Name,
        $"0x{exception.HResult:X8}",
        exception.ToString(),
        []);

    private void Request_Edited(object? sender, EventArgs e)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReady)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasIncludedRequests)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NeedsAttention)));
        Edited?.Invoke(this, EventArgs.Empty);
    }
}
