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
    private readonly string _materialFolder;
    private readonly ExcelTemplateProfileLoadResult _configuration;

    public MaterialRegistrationWindow(string materialFolder)
    {
        InitializeComponent();
        _materialFolder = Path.GetFullPath(materialFolder);
        _configuration = ExcelTemplateProfileStore.LoadDefault();

        ReferenceProfileComboBox.ItemsSource = _configuration.Profiles;
        ReferenceProfileComboBox.SelectedIndex = _configuration.Profiles.Count > 0 ? 0 : -1;
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
            InitialDirectory = _materialFolder,
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            WorkbookPathTextBox.Text = dialog.FileName;
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

            var customProfiles = _configuration.Profiles
                .Where(profile => profile.Id.StartsWith("custom-", StringComparison.OrdinalIgnoreCase))
                .Append(newProfile)
                .ToArray();
            var saved = ExcelTemplateProfileStore.SaveUserProfiles(customProfiles);
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
            message = "Excelファイルは、最初の画面で指定した教材フォルダーの中に置いてください。";
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
        if (_configuration.Profiles.Any(item =>
                string.Equals(item.DisplayName, displayName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(item.WorkbookNameKeyword, workbookKeyword, StringComparison.OrdinalIgnoreCase)))
        {
            message = "同じ教材名またはExcelファイルがすでに登録されています。";
            return false;
        }

        var keywords = KeywordsTextBox.Text
            .Split(['\r', '\n', ',', '、', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Prepend(displayName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var existingKeywords = _configuration.Profiles
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
            Id = $"custom-{DateTime.UtcNow:yyyyMMddHHmmss}-{uniqueSuffix}",
            DisplayName = displayName,
            WorkbookNameKeyword = workbookKeyword,
            MaterialNameKeywords = keywords,
            MaximumQuestionNumber = maximumRange,
            MaximumQuestionCount = maximumQuestionCount,
            HeaderTitle = null
        };
        return true;
    }

    private bool IsInsideMaterialFolder(string workbookPath)
    {
        var folderWithSeparator = Path.TrimEndingDirectorySeparator(_materialFolder)
            + Path.DirectorySeparatorChar;
        var fullWorkbookPath = Path.GetFullPath(workbookPath);
        return fullWorkbookPath.StartsWith(folderWithSeparator, StringComparison.OrdinalIgnoreCase);
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
