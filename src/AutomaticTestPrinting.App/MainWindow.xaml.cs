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
    private readonly List<RecognizedReport> _recognizedReports = [];
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
        ReportInboxFolderTextBox.Text = settings.ReportInboxFolder ?? string.Empty;
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
        InvalidateRecognitionResults();
        UpdateStatus();
    }

    private void BrowseReportInboxFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = SelectFolder(
            "Google Driveでミラーリングした未処理レポートフォルダーを選択してください");
        if (folder is null)
        {
            return;
        }

        ReportInboxFolderTextBox.Text = folder;
        _reports.Clear();
        InvalidateRecognitionResults();
        UpdateReportListState();
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

        var addedReports = ReportSelectionService.CreateDistinct(dialog.FileNames, _reports);
        foreach (var report in addedReports)
        {
            _reports.Add(report);
        }

        if (addedReports.Count > 0)
        {
            InvalidateRecognitionResults();
        }

        UpdateReportListState();
        UpdateStatus();
    }

    private async void LoadInboxReports_Click(object sender, RoutedEventArgs e)
    {
        var inboxFolder = ReportInboxFolderTextBox.Text;
        if (string.IsNullOrWhiteSpace(inboxFolder) || !Directory.Exists(inboxFolder))
        {
            MessageBox.Show(this,
                "Google Driveのレポート受け取りフォルダーを設定してください。",
                "フォルダーを確認してください",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var scanResult = ReportInboxService.Scan(inboxFolder, _reports);
        foreach (var report in scanResult.Reports)
        {
            _reports.Add(report);
        }

        if (scanResult.Reports.Count > 0)
        {
            InvalidateRecognitionResults();
        }

        UpdateReportListState();
        UpdateStatus();
        await SaveSettingsAsync();

        if (scanResult.Reports.Count == 0)
        {
            var message = scanResult.WaitingForSyncCount > 0
                ? "同期中と思われるPDFがあります。数秒待ってからもう一度読み込んでください。"
                : "新しいPDFはありませんでした。";
            MessageBox.Show(this,
                message,
                "レポート受け取りフォルダー",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        StatusText.Text = scanResult.WaitingForSyncCount > 0
            ? $"レポート {scanResult.Reports.Count}件を追加しました（同期待ち {scanResult.WaitingForSyncCount}件）"
            : $"レポート {scanResult.Reports.Count}件をGoogle Driveから追加しました";
    }

    private void ClearReports_Click(object sender, RoutedEventArgs e)
    {
        _reports.Clear();
        _recognitionResults.Clear();
        _recognizedReports.Clear();
        ConfirmResultsCheckBox.IsChecked = false;
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
        _recognizedReports.Clear();
        ConfirmResultsCheckBox.IsChecked = false;
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
                    _recognizedReports.Add(result);
                    _recognitionResults.Add(
                        RecognitionResultItem.Success(result, MaterialFolderTextBox.Text));
                }
                catch (Exception exception)
                {
                    _recognitionResults.Add(
                        RecognitionResultItem.Failure(report.FullPath, exception.Message));
                }

                UpdateRecognitionResultState();
                UpdatePdfGenerationState();
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

    private void ConfirmResultsCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        UpdatePdfGenerationState();
    }

    private async void CreatePdfsButton_Click(object sender, RoutedEventArgs e)
    {
        var requests = BuildPdfGenerationRequests();
        if (requests.Count == 0)
        {
            MessageBox.Show(this,
                "作成できる通常テストがありません。教材フォルダーと読み取り結果を確認してください。",
                "PDFを作成できません",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var expectedRequestCount = _recognizedReports.Sum(report => report.TestRequests.Count);
        if (requests.Count != expectedRequestCount)
        {
            MessageBox.Show(this,
                "未対応または入力エラーのテストがあります。青字のExcel連携表示を確認してください。",
                "PDFを作成できません",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var confirmation = MessageBox.Show(this,
            $"通常テスト {requests.Count}件の問題PDF・解答PDFを作成します。\n" +
            "プリンターへの印刷は行いません。続けますか？",
            "PDFを作成します",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        CreatePdfsButton.IsEnabled = false;
        ValidateButton.IsEnabled = false;
        ConfirmResultsCheckBox.IsEnabled = false;
        var generatedFiles = new List<string>();

        try
        {
            for (var index = 0; index < requests.Count; index++)
            {
                StatusText.Text = $"PDFを作成中：{index + 1}/{requests.Count}";
                var result = await ExcelPdfGenerationService.GenerateAsync(requests[index]);
                generatedFiles.Add(result.ProblemPdfPath);
                generatedFiles.Add(result.AnswerPdfPath);
            }

            StatusText.Text = $"PDF作成完了：{generatedFiles.Count}ファイル";
            MessageBox.Show(this,
                $"問題PDFと解答PDFを作成しました。\n\n保存先：{OutputFolderTextBox.Text}",
                "PDF作成完了",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception exception)
        {
            StatusText.Text = "PDF作成中にエラーが発生しました";
            MessageBox.Show(this,
                exception.Message,
                "PDFを作成できませんでした",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            ValidateButton.IsEnabled = true;
            ConfirmResultsCheckBox.IsEnabled = true;
            UpdatePdfGenerationState();
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
            ReportInboxFolder = EmptyToNull(ReportInboxFolderTextBox.Text),
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

    private void UpdatePdfGenerationState()
    {
        CreatePdfsButton.IsEnabled =
            ConfirmResultsCheckBox.IsChecked == true &&
            _recognizedReports.Count > 0 &&
            !_recognitionResults.Any(result => result.HasError) &&
            _recognizedReports.SelectMany(report => report.TestRequests).Any();
    }

    private void InvalidateRecognitionResults()
    {
        _recognizedReports.Clear();
        _recognitionResults.Clear();
        ConfirmResultsCheckBox.IsChecked = false;
        UpdateRecognitionResultState();
        UpdatePdfGenerationState();
    }

    private List<ExcelPdfGenerationRequest> BuildPdfGenerationRequests()
    {
        var requests = new List<ExcelPdfGenerationRequest>();
        foreach (var report in _recognizedReports)
        {
            foreach (var testRequest in report.TestRequests)
            {
                var preparation = ExcelTemplateCatalog.Prepare(
                    testRequest,
                    MaterialFolderTextBox.Text);
                if (!preparation.IsValid ||
                    preparation.Profile is null ||
                    preparation.WorkbookPath is null ||
                    preparation.StartNumber is null ||
                    preparation.EndNumber is null)
                {
                    continue;
                }

                requests.Add(new ExcelPdfGenerationRequest(
                    report.StudentName,
                    testRequest.MaterialName,
                    testRequest.QuestionCount,
                    preparation.StartNumber.Value,
                    preparation.EndNumber.Value,
                    preparation.WorkbookPath,
                    OutputFolderTextBox.Text,
                    preparation.Profile));
            }
        }

        return requests;
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
