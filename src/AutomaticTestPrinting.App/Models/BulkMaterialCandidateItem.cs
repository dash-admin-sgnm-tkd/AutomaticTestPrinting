using System.ComponentModel;
using System.Globalization;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.App.Models;

public sealed class BulkMaterialCandidateItem : INotifyPropertyChanged
{
    private bool _isSelected;
    private string _displayName;
    private string _maximumQuestionNumber;
    private ExcelMaterialFormatOption? _selectedFormat;

    public BulkMaterialCandidateItem(
        WorkbookRegistrationCandidate candidate,
        IReadOnlyList<ExcelMaterialFormatOption> matchingFormats)
    {
        Candidate = candidate;
        MatchingFormats = matchingFormats;
        _isSelected = candidate.CanRegister && !IsDerivedWordWorkbook(candidate.WorkbookFileName);
        _displayName = candidate.SuggestedDisplayName;
        _maximumQuestionNumber = candidate.SuggestedMaximumQuestionNumber > 0
            ? candidate.SuggestedMaximumQuestionNumber.ToString(CultureInfo.CurrentCulture)
            : string.Empty;
        _selectedFormat = !string.IsNullOrWhiteSpace(candidate.SuggestedFormatId)
            ? matchingFormats.FirstOrDefault(format => string.Equals(
                format.Id,
                candidate.SuggestedFormatId,
                StringComparison.OrdinalIgnoreCase))
            : matchingFormats.Count == 1
                ? matchingFormats[0]
                : null;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public WorkbookRegistrationCandidate Candidate { get; }
    public IReadOnlyList<ExcelMaterialFormatOption> MatchingFormats { get; }
    public string WorkbookFileName => Candidate.WorkbookFileName;
    public string WorkbookPath => Candidate.WorkbookPath;
    public bool CanRegister => Candidate.CanRegister;
    public string Message => IsDerivedWordWorkbook(WorkbookFileName)
        ? "派生語教材のため未選択（必要なら登録できます）"
            : SelectedFormat is null
                ? Candidate.Message
                : $"{SelectedFormat.DisplayName}を選択中";

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            var updated = CanRegister && value;
            if (_isSelected == updated)
            {
                return;
            }

            _isSelected = updated;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
        }
    }

    public string DisplayName
    {
        get => _displayName;
        set
        {
            if (_displayName == value)
            {
                return;
            }

            _displayName = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
        }
    }

    public string MaximumQuestionNumber
    {
        get => _maximumQuestionNumber;
        set
        {
            if (_maximumQuestionNumber == value)
            {
                return;
            }

            _maximumQuestionNumber = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaximumQuestionNumber)));
        }
    }

    public ExcelMaterialFormatOption? SelectedFormat
    {
        get => _selectedFormat;
        set
        {
            if (_selectedFormat == value)
            {
                return;
            }

            _selectedFormat = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedFormat)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Message)));
        }
    }

    private static bool IsDerivedWordWorkbook(string fileName) =>
        fileName.Contains("派生語", StringComparison.OrdinalIgnoreCase);
}
