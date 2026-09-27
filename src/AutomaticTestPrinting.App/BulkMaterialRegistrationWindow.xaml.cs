using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using AutomaticTestPrinting.App.Models;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.App;

public partial class BulkMaterialRegistrationWindow : Window
{
    private readonly string _materialFolder;
    private ExcelTemplateProfileLoadResult _configuration;

    public BulkMaterialRegistrationWindow(string materialFolder)
    {
        InitializeComponent();
        _materialFolder = Path.GetFullPath(materialFolder);
        _configuration = ExcelTemplateProfileStore.LoadDefault();
        DataContext = this;
    }

    public ObservableCollection<BulkMaterialCandidateItem> Candidates { get; } = [];
    public int RegisteredCount { get; private set; }
    public string RegisteredDisplayNames { get; private set; } = string.Empty;

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        await ScanAsync();
    }

    private async void Rescan_Click(object sender, RoutedEventArgs e)
    {
        await ScanAsync();
    }

    private async Task ScanAsync()
    {
        RescanButton.IsEnabled = false;
        RegisterSelectedButton.IsEnabled = false;
        Candidates.Clear();
        ScanSummaryText.Text = "教材フォルダー内のExcelを確認しています…";
        RegistrationStatusText.Text = "新しい教材を自動判定しています";

        try
        {
            _configuration = ExcelTemplateProfileStore.ReloadDefault();
            if (!_configuration.IsValid)
            {
                throw new InvalidOperationException(_configuration.Message);
            }

            var discovery = await Task.Run(() => WorkbookDiscoveryService.Discover(
                _materialFolder,
                _configuration.Profiles));
            var profilesById = _configuration.Profiles.ToDictionary(
                profile => profile.Id,
                StringComparer.OrdinalIgnoreCase);
            foreach (var candidate in discovery.Candidates)
            {
                var matchingProfiles = candidate.MatchingProfileIds
                    .Where(profilesById.ContainsKey)
                    .Select(id => profilesById[id])
                    .ToArray();
                var matchingFormats = ExcelMaterialFormatCatalog.CreateOptions(matchingProfiles);
                Candidates.Add(new BulkMaterialCandidateItem(candidate, matchingFormats));
            }

            var selectableCount = Candidates.Count(item => item.CanRegister);
            var selectedCount = Candidates.Count(item => item.IsSelected);
            var unsupportedCount = Candidates.Count - selectableCount;
            ScanSummaryText.Text = Candidates.Count == 0
                ? $"新しく登録するExcelはありません。登録済み：{discovery.RegisteredWorkbookCount}件"
                : $"新規候補：{selectableCount}件、要個別確認：{unsupportedCount}件、" +
                  $"登録済み：{discovery.RegisteredWorkbookCount}件";
            RegistrationStatusText.Text = selectableCount == 0
                ? "自動登録できる新しい教材はありません"
                : $"{selectedCount}件を選択中です。教材名と形式を確認してください";
            RegisterSelectedButton.IsEnabled = selectedCount > 0;
        }
        catch (Exception exception)
        {
            ScanSummaryText.Text = $"教材を確認できませんでした：{exception.Message}";
            RegistrationStatusText.Text = "教材フォルダーを確認してください";
            MessageBox.Show(this,
                exception.Message,
                "教材をスキャンできませんでした",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            RescanButton.IsEnabled = true;
        }
    }

    private void RegisterSelected_Click(object sender, RoutedEventArgs e)
    {
        CandidatesGrid.CommitEdit();
        CandidatesGrid.CommitEdit();
        var selected = Candidates.Where(item => item.IsSelected).ToArray();
        if (selected.Length == 0)
        {
            MessageBox.Show(this,
                "登録する教材にチェックを入れてください。",
                "教材が選択されていません",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!TryBuildProfiles(selected, out var newProfiles, out var message))
        {
            MessageBox.Show(this,
                message,
                "登録内容を確認してください",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var confirmation = MessageBox.Show(this,
            $"選択した教材 {newProfiles.Length}件をまとめて登録します。続けますか？",
            "教材をまとめて登録",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        RegisterSelectedButton.IsEnabled = false;
        RegistrationStatusText.Text = $"教材 {newProfiles.Length}件を登録しています…";
        var customProfiles = _configuration.Profiles
            .Where(profile => profile.Id.StartsWith("custom-", StringComparison.OrdinalIgnoreCase))
            .Concat(newProfiles)
            .ToArray();
        var saved = ExcelTemplateProfileStore.SaveUserProfiles(customProfiles);
        if (!saved.IsValid)
        {
            RegisterSelectedButton.IsEnabled = true;
            RegistrationStatusText.Text = "教材を登録できませんでした";
            MessageBox.Show(this,
                saved.Message,
                "教材を登録できませんでした",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        RegisteredCount = newProfiles.Length;
        RegisteredDisplayNames = string.Join("、", newProfiles.Select(profile => profile.DisplayName));
        MessageBox.Show(this,
            $"教材を{RegisteredCount}件登録しました。\n\n{RegisteredDisplayNames}",
            "まとめて登録しました",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
        DialogResult = true;
    }

    private bool TryBuildProfiles(
        BulkMaterialCandidateItem[] selected,
        out ExcelTemplateProfile[] profiles,
        out string message)
    {
        profiles = [];
        message = string.Empty;
        var existingNames = _configuration.Profiles
            .Select(profile => Normalize(profile.DisplayName))
            .ToHashSet(StringComparer.Ordinal);
        var batchNames = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<ExcelTemplateProfile>(selected.Length);

        foreach (var item in selected)
        {
            var displayName = item.DisplayName.Trim();
            if (displayName.Length == 0)
            {
                message = $"「{item.WorkbookFileName}」の教材名を入力してください。";
                return false;
            }

            var normalizedName = Normalize(displayName);
            if (existingNames.Contains(normalizedName) || !batchNames.Add(normalizedName))
            {
                message = $"教材名「{displayName}」が重複しています。";
                return false;
            }

            if (item.SelectedFormat is null)
            {
                message = $"「{item.WorkbookFileName}」の形式を選んでください。";
                return false;
            }

            if (!int.TryParse(
                    item.MaximumQuestionNumber,
                    NumberStyles.Integer,
                    CultureInfo.CurrentCulture,
                    out var maximumQuestionNumber) ||
                maximumQuestionNumber < 1)
            {
                message = $"「{displayName}」の最大番号を1以上の数字で入力してください。";
                return false;
            }

            var workbookStem = Path.GetFileNameWithoutExtension(item.WorkbookPath);
            var keywords = new[] { displayName, item.Candidate.SuggestedDisplayName }
                .Where(keyword => !string.IsNullOrWhiteSpace(keyword))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var uniqueSuffix = Guid.NewGuid().ToString("N")[..8];
            result.Add(item.SelectedFormat.TemplateProfile with
            {
                Id = $"custom-{DateTime.UtcNow:yyyyMMddHHmmss}-{uniqueSuffix}",
                DisplayName = displayName,
                WorkbookNameKeyword = workbookStem,
                WorkbookNameExcludedKeywords = [],
                MaterialNameKeywords = keywords,
                MaximumQuestionNumber = maximumQuestionNumber,
                HeaderTitle = null
            });
        }

        profiles = result.ToArray();
        return true;
    }

    private void OpenManualRegistration_Click(object sender, RoutedEventArgs e)
    {
        var window = new MaterialRegistrationWindow(_materialFolder)
        {
            Owner = this
        };
        if (window.ShowDialog() != true)
        {
            return;
        }

        RegisteredCount = 1;
        RegisteredDisplayNames = window.RegisteredDisplayName;
        DialogResult = true;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private static string Normalize(string value)
    {
        var normalized = value.Normalize(NormalizationForm.FormKC);
        return string.Concat(normalized
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant));
    }
}
