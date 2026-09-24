using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using AutomaticTestPrinting.App.Models;
using AutomaticTestPrinting.App.Services;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using Microsoft.Win32;

namespace AutomaticTestPrinting.App;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<ReportFile> _reports = [];
    private readonly ObservableCollection<RecognitionResultItem> _recognitionResults = [];
    private readonly JsonSettingsStore _settingsStore;

    public MainWindow()
    {
        InitializeComponent();

        var settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AutomaticTestPrinting",
            "settings.json");

        _settingsStore = new JsonSettingsStore(settingsPath);
        DataContext = new { Reports = _reports, RecognitionResults = _recognitionResults };
        UpdateReportListState();
        UpdateRecognitionResultState();
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
        _recognitionResults.Clear();
        UpdateReportListState();
        UpdateRecognitionResultState();
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
        ValidateButton.IsEnabled = false;
        _recognitionResults.Clear();
        UpdateRecognitionResultState();

        try
        {
            var service = new WindowsReportOcrService();
            var progress = new Progress<string>(message => StatusText.Text = message);

            foreach (var report in _reports)
            {
                try
                {
                    var result = await service.RecognizeAsync(report.FullPath, progress);
                    _recognitionResults.Add(RecognitionResultItem.Success(result));
                }
                catch (Exception exception)
                {
                    _recognitionResults.Add(
                        RecognitionResultItem.Failure(report.FullPath, exception.Message));
                }

                UpdateRecognitionResultState();
            }

            var failedCount = _recognitionResults.Count(result => result.HasError);
            StatusText.Text = failedCount == 0
                ? $"読み取り完了：{_recognitionResults.Count}件"
                : $"読み取り完了：成功 {_recognitionResults.Count - failedCount}件、要確認 {failedCount}件";
        }
        finally
        {
            ValidateButton.IsEnabled = true;
        }
    }

    private void OpenSourceReport_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string sourcePath } || !File.Exists(sourcePath))
        {
            MessageBox.Show(this, "原本PDFが見つかりません。", "ファイルを開けません",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Process.Start(new ProcessStartInfo(sourcePath) { UseShellExecute = true });
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

    private void UpdateRecognitionResultState()
    {
        RecognitionResultsSection.Visibility = _recognitionResults.Count == 0
            ? Visibility.Collapsed
            : Visibility.Visible;
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
