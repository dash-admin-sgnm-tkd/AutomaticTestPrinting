using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using Microsoft.Win32;

namespace AutomaticTestPrinting.App;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<ReportFile> _reports = [];
    private readonly JsonSettingsStore _settingsStore;

    public MainWindow()
    {
        InitializeComponent();

        var settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AutomaticTestPrinting",
            "settings.json");

        _settingsStore = new JsonSettingsStore(settingsPath);
        DataContext = new { Reports = _reports };
        UpdateReportListState();
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        var settings = await _settingsStore.LoadAsync();
        MaterialFolderTextBox.Text = settings.MaterialFolder ?? string.Empty;
        OutputFolderTextBox.Text = settings.OutputFolder ?? string.Empty;
        UpdateStatus();
    }

    private async void Window_Closing(object? sender, CancelEventArgs e)
    {
        await SaveSettingsAsync();
    }

    private void BrowseMaterialFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("Excel・PDF教材を保存しているフォルダーを選択してください");
        if (folder is null)
        {
            return;
        }

        MaterialFolderTextBox.Text = folder;
        UpdateStatus();
    }

    private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder("作成したテストを保存するフォルダーを選択してください");
        if (folder is null)
        {
            return;
        }

        OutputFolderTextBox.Text = folder;
        UpdateStatus();
    }

    private void AddReports_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "特訓レポートを選択",
            Filter = "PDFファイル (*.pdf)|*.pdf",
            Multiselect = true,
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        foreach (var report in ReportSelectionService.CreateDistinct(dialog.FileNames, _reports))
        {
            _reports.Add(report);
        }

        UpdateReportListState();
        UpdateStatus();
    }

    private void ClearReports_Click(object sender, RoutedEventArgs e)
    {
        _reports.Clear();
        UpdateReportListState();
        UpdateStatus();
    }

    private async void ValidateButton_Click(object sender, RoutedEventArgs e)
    {
        var validation = SetupValidator.Validate(
            MaterialFolderTextBox.Text,
            OutputFolderTextBox.Text,
            _reports);

        if (!validation.IsValid)
        {
            MessageBox.Show(this, validation.Message, "設定を確認してください",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        await SaveSettingsAsync();
        StatusText.Text = $"準備完了：{_reports.Count}件のレポートを読み取れます";

        MessageBox.Show(
            this,
            $"教材と出力先を確認しました。\nレポート {_reports.Count}件が読み取り対象です。\n\n次の開発段階でOCR処理を接続します。",
            "読み取り準備が整いました",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private static string? SelectFolder(string title)
    {
        var dialog = new OpenFolderDialog
        {
            Title = title,
            Multiselect = false
        };

        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }

    private async Task SaveSettingsAsync()
    {
        var settings = new AppSettings
        {
            MaterialFolder = EmptyToNull(MaterialFolderTextBox.Text),
            OutputFolder = EmptyToNull(OutputFolderTextBox.Text)
        };

        await _settingsStore.SaveAsync(settings);
    }

    private void UpdateReportListState()
    {
        EmptyReportsText.Visibility = _reports.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        ReportsList.Visibility = _reports.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
    }

    private void UpdateStatus()
    {
        var foldersSelected = !string.IsNullOrWhiteSpace(MaterialFolderTextBox.Text)
            && !string.IsNullOrWhiteSpace(OutputFolderTextBox.Text);

        StatusText.Text = foldersSelected && _reports.Count > 0
            ? $"準備確認できます：レポート {_reports.Count}件"
            : "設定を確認してください";
    }

    private static string? EmptyToNull(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }
}
