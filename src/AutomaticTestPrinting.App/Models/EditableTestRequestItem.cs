using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.App.Models;

public sealed class EditableTestRequestItem : INotifyPropertyChanged
{
    private readonly IReadOnlyCollection<string> _materialFolders;
    private string _materialName;
    private string _startNumber;
    private string _endNumber;
    private string _questionCount;
    private string _validationMessage = string.Empty;
    private bool _isValid;
    private bool _isIncluded = true;
    private string? _suggestionMaterialName;

    public EditableTestRequestItem(
        NormalTestRequestCandidate request,
        int displayNumber,
        IReadOnlyCollection<string> materialFolders)
    {
        _materialFolders = materialFolders;
        DisplayNumber = displayNumber;
        PageNumber = request.PageNumber;
        RequiresReview = request.RequiresReview;
        OriginalMaterialName = request.MaterialName;
        _materialName = ConfirmedReportMaterialNames.Resolve(request.MaterialName);
        (_startNumber, _endNumber) = SplitRange(request.Range);
        _questionCount = request.QuestionCount > 0 ? request.QuestionCount.ToString(CultureInfo.CurrentCulture) : string.Empty;
        Revalidate();
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? Edited;

    public int DisplayNumber { get; }
    public string RequestHeading => PdfMaterialProfileStore.NormalizeName(MaterialName)
        .StartsWith("段階突破", StringComparison.Ordinal) ? "段階突破テストを作成する" : $"テスト {DisplayNumber} を作成する";
    public int PageNumber { get; }
    public bool RequiresReview { get; }
    public string OriginalMaterialName { get; }
    public IReadOnlyList<string> MaterialSuggestions { get; private set; } = [];
    public bool HasMaterialSuggestions => MaterialSuggestions.Count > 0;
    public string? SelectedMaterialSuggestion { get; set; }
    public IReadOnlyList<string> WorkbookCandidates { get; private set; } = [];
    public ExcelTemplateProfile? WorkbookProfile { get; private set; }
    public bool CanChooseWorkbook => WorkbookCandidates.Count > 0;
    public bool CanEditWorkbook => WorkbookProfile is not null && !IsValid;

    public string MaterialName
    {
        get => _materialName;
        set => SetField(ref _materialName, value);
    }

    public string StartNumber
    {
        get => _startNumber;
        set => SetField(ref _startNumber, value);
    }

    public string EndNumber
    {
        get => _endNumber;
        set => SetField(ref _endNumber, value);
    }

    public string QuestionCount
    {
        get => _questionCount;
        set => SetField(ref _questionCount, value);
    }

    public string ValidationMessage
    {
        get => _validationMessage;
        private set => SetValue(ref _validationMessage, value);
    }

    public bool IsValid
    {
        get => _isValid;
        private set
        {
            if (_isValid == value)
            {
                return;
            }

            _isValid = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsValid)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanEditWorkbook)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsReady)));
        }
    }

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
            Edited?.Invoke(this, EventArgs.Empty);
        }
    }

    public bool IsReady => !IsIncluded || IsValid;

    public NormalTestRequestCandidate? BuildCandidate()
    {
        if (!int.TryParse(QuestionCount.Trim(), NumberStyles.None, CultureInfo.CurrentCulture, out var count) ||
            count < 1)
        {
            return null;
        }

        return new NormalTestRequestCandidate(
            MaterialName.Trim(),
            $"{StartNumber.Trim()}-{EndNumber.Trim()}",
            count,
            PageNumber,
            RequiresReview);
    }

    private void UpdateMaterialSuggestions(bool force)
    {
        if (!force && _suggestionMaterialName == MaterialName) return;
        _suggestionMaterialName = MaterialName;
        try
        {
            var materials = SharedMaterialCatalog.Load().Items.Select(item => (Name: item.DisplayName, Aliases: (IEnumerable<string>)item.Aliases))
                .Concat(ExcelTemplateProfileStore.LoadDefault().Profiles.Select(profile => (Name: profile.DisplayName, Aliases: (IEnumerable<string>)profile.MaterialNameKeywords)))
                .Concat(PdfMaterialProfileStore.Load().Select(profile => (Name: profile.DisplayName, Aliases: (IEnumerable<string>)profile.Aliases)));
            MaterialSuggestions = MaterialNameSuggestionService.Suggest(MaterialName, materials);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ArgumentException)
        {
            // Optional hints must not replace the specific configuration error shown by validation.
            MaterialSuggestions = [];
        }
        SelectedMaterialSuggestion = null;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaterialSuggestions)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasMaterialSuggestions)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedMaterialSuggestion)));
    }

    private void Revalidate(bool refreshSuggestions = false)
    {
        UpdateMaterialSuggestions(refreshSuggestions);
        WorkbookCandidates = [];
        WorkbookProfile = null;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanEditWorkbook)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanChooseWorkbook)));
        var candidate = BuildCandidate();
        if (candidate is null || string.IsNullOrWhiteSpace(candidate.MaterialName))
        {
            IsValid = false;
            ValidationMessage = "教材名・範囲・問題数を原本と照合して入力してください。空欄は依頼なし・読み取り漏れの可能性があります。";
            return;
        }

        var pdf = PdfMaterialCatalog.Prepare(candidate, _materialFolders);
        if (pdf.IsSupported)
        {
            IsValid = pdf.IsValid;
            ValidationMessage = pdf.Message;
        }
        else
        {
            var preparation = ExcelTemplateCatalog.Prepare(candidate, _materialFolders);
            ApplyExcelPreparation(preparation);
            if (!preparation.IsSupported && MaterialSuggestions.Count > 0)
                ValidationMessage = "教材名が登録と一致しません（誤読・略称の可能性）。原本を見て下の修正候補を確認してください。";
        }
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanChooseWorkbook)));
    }

    public void RefreshValidation()
    {
        Revalidate(refreshSuggestions: true);
        Edited?.Invoke(this, EventArgs.Empty);
    }

    public void ApplyExcelPreparation(ExcelRequestPreparation preparation)
    {
        WorkbookProfile = preparation.Profile;
        WorkbookCandidates = preparation.WorkbookCandidates ?? [];
        IsValid = preparation.IsValid;
        ValidationMessage = preparation.Message;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanEditWorkbook)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CanChooseWorkbook)));
    }

    private void SetField(ref string field, string value, [CallerMemberName] string? propertyName = null)
    {
        if (string.Equals(field, value, StringComparison.Ordinal))
        {
            return;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        if (propertyName == nameof(MaterialName))
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RequestHeading)));
        Revalidate();
        Edited?.Invoke(this, EventArgs.Empty);
    }

    private void SetValue<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static (string Start, string End) SplitRange(string range)
    {
        var separatorIndex = range.IndexOf('-');
        return separatorIndex > 0
            ? (range[..separatorIndex].Trim(), range[(separatorIndex + 1)..].Trim())
            : (range.Trim(), string.Empty);
    }
}
