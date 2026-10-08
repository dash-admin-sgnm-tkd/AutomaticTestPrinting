using System.IO;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using AutomaticTestPrinting.App.Services;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using Microsoft.Win32;

namespace AutomaticTestPrinting.App;

public partial class MaterialRegistrationWindow : Window
{
    private readonly string[] _materialFolders;
    private readonly ExcelTemplateProfileLoadResult _configuration;
    private ExcelTemplateProfile? _editingProfile;

    public MaterialRegistrationWindow(IReadOnlyCollection<string> materialFolders, SharedMaterialEntry? catalogEntry = null)
    {
        InitializeComponent();
        _materialFolders = materialFolders
            .Where(Directory.Exists)
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        _configuration = ExcelTemplateProfileStore.ReloadDefault();

        ReferenceProfileComboBox.ItemsSource = _configuration.Profiles;
        ReferenceProfileComboBox.SelectedIndex = _configuration.Profiles.Count > 0 ? 0 : -1;
        if (catalogEntry is not null)
        {
            ReferenceProfileComboBox.SelectedIndex = -1;
            DisplayNameTextBox.Text = catalogEntry.DisplayName;
            KeywordsTextBox.Text = string.Join(Environment.NewLine, catalogEntry.Aliases);
            MaximumRangeTextBox.Clear();
            MaximumQuestionCountTextBox.Clear();
            RegistrationStatusText.Text = "カタログの教材名を入力済みです。使用ファイル・形式・上限を確認してください。";
            var registered = ExcelMaterialEditingService.FindRegistered(_configuration.Profiles,
                catalogEntry.FileNames, catalogEntry.DisplayName);
            if (registered.Count == 1) BeginEditing(registered[0]);
            else if (registered.Count > 1)
            {
                RegistrationStatusText.Text = "登録先が複数あります。重複する設定の確認が必要です：" +
                    string.Join(" / ", registered.Select(p => p.DisplayName));
                RegisterButton.IsEnabled = false;
            }
            var fileName = _editingProfile?.WorkbookFileName;
            if (string.IsNullOrWhiteSpace(fileName))
                fileName = catalogEntry.FileNames.Count > 0 ? catalogEntry.FileNames[0] : null;
            var paths = _materialFolders.SelectMany(folder => new[] {
                Path.Combine(folder, catalogEntry.SourceRelativePath.Replace('/', Path.DirectorySeparatorChar)),
                Path.Combine(folder, fileName ?? "") })
                .Where(path => File.Exists(path) && IsInsideMaterialFolder(path) &&
                    string.Equals(Path.GetFileName(path), fileName, StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (paths.Length == 1) WorkbookPathTextBox.Text = paths[0];
        }
    }

    private void BeginEditing(ExcelTemplateProfile profile)
    {
        _editingProfile = profile;
        Title = "登録済み教材を編集";
        HeadingText.Text = "登録済み教材を編集";
        DescriptionText.Text = "この教材の現在の設定を表示しています。変更後は「確認して保存」を押してください。他の教材は変更しません。";
        ReferenceProfileComboBox.SelectedItem = profile;
        DisplayNameTextBox.Text = profile.DisplayName;
        KeywordsTextBox.Text = string.Join(Environment.NewLine, profile.MaterialNameKeywords);
        MaximumRangeTextBox.Text = profile.MaximumQuestionNumber.ToString(CultureInfo.CurrentCulture);
        MaximumQuestionCountTextBox.Text = profile.MaximumQuestionCount.ToString(CultureInfo.CurrentCulture);
        RegisterButton.Content = "確認して保存";
        RegistrationStatusText.Text = $"「{profile.DisplayName}」の設定を編集中。使用するExcelファイルを確認してください。";
    }

    public string RegisteredDisplayName { get; private set; } = string.Empty;

    private void ReferenceProfileComboBox_SelectionChanged(
        object sender,
        SelectionChangedEventArgs e)
    {
        if (ReferenceProfileComboBox.SelectedItem is not ExcelTemplateProfile profile)
        {
            return;
        }

        MaximumRangeTextBox.Text = profile.MaximumQuestionNumber.ToString(CultureInfo.CurrentCulture);
        MaximumQuestionCountTextBox.Text = profile.MaximumQuestionCount.ToString(CultureInfo.CurrentCulture);
    }

    private void BrowseWorkbook_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "登録する教材のExcelファイルを選択",
            Filter = "マクロ有効Excelファイル (*.xlsm)|*.xlsm",
            InitialDirectory = _materialFolders.Length > 0 ? _materialFolders[0] : null,
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            WorkbookPathTextBox.Text = dialog.FileName;
            if (_editingProfile is null)
            {
                var registered = ExcelMaterialEditingService.FindRegistered(_configuration.Profiles,
                    [Path.GetFileName(dialog.FileName)], "");
                if (registered.Count == 1 && MessageBox.Show(this,
                    $"このExcelは「{registered[0].DisplayName}」として登録済みです。現在の設定を開いて編集しますか？",
                    "登録済み教材", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    BeginEditing(registered[0]);
            }
        }
    }

    private async void Register_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildProfile(out var newProfile, out var message) || newProfile is null)
        {
            MessageBox.Show(this, message, "入力内容を確認してください",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        RegisterButton.IsEnabled = false;
        RegistrationStatusText.Text = "Excelの形式を確認しています…";
        try
        {
            var referenceProfile = (ExcelTemplateProfile)ReferenceProfileComboBox.SelectedItem;
            var validation = await ExcelMaterialRegistrationValidator.ValidateAsync(
                WorkbookPathTextBox.Text,
                referenceProfile);
            if (!validation.IsValid)
            {
                MessageBox.Show(this, validation.Message, "Excelの形式を確認してください",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                RegistrationStatusText.Text = "形式が近い教材を選び直してください";
                return;
            }

            var saved = ExcelTemplateProfileStore.SaveUserProfileUpdates([newProfile]);
            if (!saved.IsValid)
            {
                MessageBox.Show(this, saved.Message, "教材を登録できませんでした",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                RegistrationStatusText.Text = "保存できませんでした";
                return;
            }

            RegisteredDisplayName = newProfile.DisplayName;
            DialogResult = true;
        }
        finally
        {
            RegisterButton.IsEnabled = true;
        }
    }

    private bool TryBuildProfile(
        out ExcelTemplateProfile? profile,
        out string message)
    {
        profile = null;
        message = string.Empty;

        if (ReferenceProfileComboBox.SelectedItem is not ExcelTemplateProfile referenceProfile)
        {
            message = "形式が近い教材を選んでください。";
            return false;
        }

        var displayName = DisplayNameTextBox.Text.Trim();
        if (displayName.Length == 0)
        {
            message = "教材名を入力してください。";
            return false;
        }

        var workbookPath = WorkbookPathTextBox.Text.Trim();
        if (!File.Exists(workbookPath) ||
            !string.Equals(Path.GetExtension(workbookPath), ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            message = "教材の .xlsm ファイルを選択してください。";
            return false;
        }

        if (!IsInsideMaterialFolder(workbookPath))
        {
            message = "Excelファイルは、最初の画面で追加した教材フォルダーのいずれかに置いてください。";
            return false;
        }

        if (!int.TryParse(MaximumRangeTextBox.Text, out var maximumRange) || maximumRange < 1)
        {
            message = "最大の問題番号を1以上の数字で入力してください。";
            return false;
        }

        if (!int.TryParse(MaximumQuestionCountTextBox.Text, out var maximumQuestionCount) ||
            maximumQuestionCount < 1)
        {
            message = "1回の最大問題数を1以上の数字で入力してください。";
            return false;
        }

        var workbookKeyword = Path.GetFileNameWithoutExtension(workbookPath);
        var otherProfiles = _configuration.Profiles.Where(item =>
            !string.Equals(item.Id, _editingProfile?.Id, StringComparison.OrdinalIgnoreCase)).ToArray();
        var duplicate = otherProfiles.FirstOrDefault(item =>
                string.Equals(item.DisplayName, displayName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(item.WorkbookNameKeyword, workbookKeyword, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(item.WorkbookFileName, Path.GetFileName(workbookPath), StringComparison.OrdinalIgnoreCase));
        if (duplicate is not null)
        {
            message = $"「{duplicate.DisplayName}」で同じ教材名またはExcelファイルが登録済みです。カタログからその教材を選び「選択教材の登録・修正」で編集してください。";
            return false;
        }

        var keywords = KeywordsTextBox.Text
            .Split(['\r', '\n', ',', '、', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Prepend(displayName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var existingKeywords = otherProfiles
            .SelectMany(item => item.MaterialNameKeywords)
            .Select(NormalizeForComparison)
            .ToHashSet(StringComparer.Ordinal);
        var duplicateKeyword = keywords.FirstOrDefault(keyword =>
            existingKeywords.Contains(NormalizeForComparison(keyword)));
        if (duplicateKeyword is not null)
        {
            message = $"「{duplicateKeyword}」は別の教材ですでに使われている名前です。";
            return false;
        }

        var uniqueSuffix = Guid.NewGuid().ToString("N")[..6];

        profile = referenceProfile with
        {
            Id = _editingProfile?.Id ?? $"custom-{DateTime.UtcNow:yyyyMMddHHmmss}-{uniqueSuffix}",
            DisplayName = displayName,
            WorkbookNameKeyword = workbookKeyword,
            WorkbookFileName = Path.GetFileName(workbookPath),
            WorkbookNameExcludedKeywords = _editingProfile is not null &&
                string.Equals(_editingProfile.WorkbookNameKeyword, workbookKeyword, StringComparison.OrdinalIgnoreCase)
                ? _editingProfile.WorkbookNameExcludedKeywords : [],
            MaterialNameKeywords = keywords,
            MaximumQuestionNumber = maximumRange,
            MaximumQuestionCount = maximumQuestionCount,
            HeaderTitle = _editingProfile is not null && referenceProfile.Id == _editingProfile.Id
                ? _editingProfile.HeaderTitle : null
        };
        return true;
    }

    private bool IsInsideMaterialFolder(string workbookPath)
    {
        var fullWorkbookPath = Path.GetFullPath(workbookPath);
        return _materialFolders.Any(folder =>
        {
            var folderWithSeparator = Path.TrimEndingDirectorySeparator(folder)
                + Path.DirectorySeparatorChar;
            return fullWorkbookPath.StartsWith(folderWithSeparator, StringComparison.OrdinalIgnoreCase);
        });
    }

    private static string NormalizeForComparison(string value)
    {
        var normalized = value.Normalize(NormalizationForm.FormKC);
        return string.Concat(normalized
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant));
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
