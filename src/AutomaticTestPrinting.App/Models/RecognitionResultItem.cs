using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.App.Models;

public sealed class RecognitionResultItem : INotifyPropertyChanged
{
    private string _studentName;

    private RecognitionResultItem(
        string sourcePath,
        string studentName,
        string normalTestSummary,
        string stageTestSummary,
        string orientationSummary,
        bool hasError,
        IEnumerable<EditableTestRequestItem> requests)
    {
        SourcePath = sourcePath;
        FileName = Path.GetFileName(sourcePath);
        _studentName = studentName;
        NormalTestSummary = normalTestSummary;
        StageTestSummary = stageTestSummary;
        OrientationSummary = orientationSummary;
        HasError = hasError;
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
    public ObservableCollection<EditableTestRequestItem> Requests { get; }

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
        !HasError &&
        !string.IsNullOrWhiteSpace(StudentName) &&
        Requests.Count > 0 &&
        Requests.All(request => request.IsValid);

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
            requests);
    }

    public static RecognitionResultItem Failure(string sourcePath, string message) => new(
        sourcePath,
        "読み取りに失敗しました",
        message,
        "PDFを開けるか確認してください",
        string.Empty,
        true,
        []);

    private void Request_Edited(object? sender, EventArgs e)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReady)));
        Edited?.Invoke(this, EventArgs.Empty);
    }
}
