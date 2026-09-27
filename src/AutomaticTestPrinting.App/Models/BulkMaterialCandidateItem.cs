using System.ComponentModel;
using System.Globalization;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.App.Models;

public sealed class BulkMaterialCandidateItem : INotifyPropertyChanged
{
    private bool _isSelected;
    private string _displayName;
    private string _maximumQuestionNumber;
    private ExcelTemplateProfile? _selectedProfile;

    public BulkMaterialCandidateItem(
        WorkbookRegistrationCandidate candidate,
        IReadOnlyList<ExcelTemplateProfile> matchingProfiles)
    {
        Candidate = candidate;
        MatchingProfiles = matchingProfiles;
        _isSelected = candidate.CanRegister && !IsDerivedWordWorkbook(candidate.WorkbookFileName);
        _displayName = candidate.SuggestedDisplayName;
        _maximumQuestionNumber = candidate.SuggestedMaximumQuestionNumber > 0
            ? candidate.SuggestedMaximumQuestionNumber.ToString(CultureInfo.CurrentCulture)
            : string.Empty;
        _selectedProfile = matchingProfiles.Count > 0 ? matchingProfiles[0] : null;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public WorkbookRegistrationCandidate Candidate { get; }
    public IReadOnlyList<ExcelTemplateProfile> MatchingProfiles { get; }
    public string WorkbookFileName => Candidate.WorkbookFileName;
    public string WorkbookPath => Candidate.WorkbookPath;
    public bool CanRegister => Candidate.CanRegister;
    public string Message => IsDerivedWordWorkbook(WorkbookFileName)
        ? "派生語教材のため未選択（必要なら登録できます）"
        : Candidate.Message;

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

    public ExcelTemplateProfile? SelectedProfile
    {
        get => _selectedProfile;
        set
        {
            if (_selectedProfile == value)
            {
                return;
            }

            _selectedProfile = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedProfile)));
        }
    }

    private static bool IsDerivedWordWorkbook(string fileName) =>
        fileName.Contains("派生語", StringComparison.OrdinalIgnoreCase);
}
