using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.App.Models;

public sealed class RecognitionResultItem : INotifyPropertyChanged
{
    private string _studentName;
    private bool _isIncluded = true;
    public bool IsDetailsExpanded { get; set; }

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
            request.PropertyChanged += Request_PropertyChanged;
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? Edited;

    public string SourcePath { get; }
    public string FileName { get; }
    public string NormalTestSummary { get; }
    public string NormalTestReadWarning { get; private init; } = string.Empty;
    public bool HasReadingWarning => NormalTestReadWarning.Length > 0;
    public string StageTestSummary { get; }
    public bool ShowStageTestSummary { get; private init; } = true;
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
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReviewStatus)));
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
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HeaderName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NeedsAttention)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReviewStatus)));
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
        (HasError || HasReadingWarning || !IsReady);

    public string HeaderName => HasError || string.IsNullOrWhiteSpace(StudentName) ? FileName : StudentName;

    public string ReviewStatus
    {
        get
        {
            if (!IsIncluded) return "対象外";
            if (HasError) return "読み取りエラー";
            if (NeedsAttention) return "要確認";
            if (HasIncludedRequests) return "作成可能";
            return Requests.Count == 0 ? "依頼なし" : "対象外";
        }
    }

    public static RecognitionResultItem Success(
        RecognizedReport report,
        IReadOnlyCollection<string> materialFolders)
    {
        var requests = report.TestRequests
            .Select((request, index) => new EditableTestRequestItem(request, index + 1, materialFolders))
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
            requests) { ShowStageTestSummary = report.StageTestRequest is null, NormalTestReadWarning = report.NormalTestReadWarning };
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
        NotifyReviewStatusChanged();
        Edited?.Invoke(this, EventArgs.Empty);
    }

    private void Request_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(EditableTestRequestItem.IsValid))
        {
            NotifyReviewStatusChanged();
        }
    }

    private void NotifyReviewStatusChanged()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReady)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasIncludedRequests)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NeedsAttention)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReviewStatus)));
    }
}
