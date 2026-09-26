using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.App.Models;

public sealed class EditableTestRequestItem : INotifyPropertyChanged
{
    private readonly string? _materialFolder;
    private string _materialName;
    private string _startNumber;
    private string _endNumber;
    private string _questionCount;
    private string _validationMessage = string.Empty;
    private bool _isValid;

    public EditableTestRequestItem(
        NormalTestRequestCandidate request,
        int displayNumber,
        string? materialFolder)
    {
        _materialFolder = materialFolder;
        DisplayNumber = displayNumber;
        PageNumber = request.PageNumber;
        RequiresReview = request.RequiresReview;
        _materialName = request.MaterialName;
        (_startNumber, _endNumber) = SplitRange(request.Range);
        _questionCount = request.QuestionCount.ToString(CultureInfo.CurrentCulture);
        Revalidate();
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? Edited;

    public int DisplayNumber { get; }
    public int PageNumber { get; }
    public bool RequiresReview { get; }

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
        private set => SetValue(ref _isValid, value);
    }

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

    private void Revalidate()
    {
        var candidate = BuildCandidate();
        if (candidate is null || string.IsNullOrWhiteSpace(candidate.MaterialName))
        {
            IsValid = false;
            ValidationMessage = "教材名・範囲・問題数を入力してください。";
            return;
        }

        var preparation = ExcelTemplateCatalog.Prepare(candidate, _materialFolder);
        IsValid = preparation.IsValid;
        ValidationMessage = preparation.Message;
    }

    private void SetField(ref string field, string value, [CallerMemberName] string? propertyName = null)
    {
        if (string.Equals(field, value, StringComparison.Ordinal))
        {
            return;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
