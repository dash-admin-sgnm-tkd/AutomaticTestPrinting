using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Data;
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
    private readonly ICollectionView _recognitionResultsView;
    private readonly JsonSettingsStore _settingsStore;
    private bool _showOnlyNeedsAttention;

    public MainWindow()
    {
        InitializeComponent();

        var settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AutomaticTestPrinting",
            "settings.json");

        _settingsStore = new JsonSettingsStore(settingsPath);
        _recognitionResultsView = CollectionViewSource.GetDefaultView(_recognitionResults);
        DataContext = new { Reports = _reports, RecognitionResultsView = _recognitionResultsView };
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

    private void RegisterMaterial_Click(object sender, RoutedEventArgs e)
    {
        var materialFolder = MaterialFolderTextBox.Text;
        if (string.IsNullOrWhiteSpace(materialFolder) || !Directory.Exists(materialFolder))
        {
            MessageBox.Show(this,
                "先に教材フォルダーを設定してください。",
                "教材フォルダーが必要です",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var registrationWindow = new MaterialRegistrationWindow(materialFolder)
        {
            Owner = this
        };
        if (registrationWindow.ShowDialog() != true)
        {
            return;
        }

        InvalidateRecognitionResults();
        StatusText.Text = $"教材「{registrationWindow.RegisteredDisplayName}」を登録しました";
        MessageBox.Show(this,
            $"教材「{registrationWindow.RegisteredDisplayName}」を登録しました。\n" +
            "次回からレポートの読み取り候補として使用されます。",
            "教材を登録しました",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
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
        ConfirmResultsCheckBox.IsChecked = false;
        UpdateRecognitionResultState();

        try
        {
            var reports = _reports.ToArray();
            IProgress<string> progress = new Progress<string>(message => StatusText.Text = message);
            using var concurrencyGate = new SemaphoreSlim(2);
            var completedCount = 0;
            var recognitionTasks = reports.Select(async report =>
            {
                await concurrencyGate.WaitAsync();
                try
                {
                    var service = new WindowsReportOcrService();
                    var recognized = await service.RecognizeAsync(report.FullPath, progress);
                    return RecognitionResultItem.Success(recognized, MaterialFolderTextBox.Text);
                }
                catch (Exception exception)
                {
                    return RecognitionResultItem.Failure(report.FullPath, exception.Message);
                }
                finally
                {
                    concurrencyGate.Release();
                    var current = Interlocked.Increment(ref completedCount);
                    progress.Report($"読み取り中：{current}/{reports.Length}件完了");
                }
            }).ToArray();

            var recognitionResults = await Task.WhenAll(recognitionTasks);
            foreach (var result in recognitionResults)
            {
                AddRecognitionResult(result);
            }

            UpdateRecognitionResultState();
            UpdatePdfGenerationState();

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

    private void SkipInvalidRequests_Click(object sender, RoutedEventArgs e)
    {
        var skippedCount = 0;
        ConfirmResultsCheckBox.IsChecked = false;
        foreach (var result in _recognitionResults)
        {
            if (result.HasError)
            {
                if (result.IsIncluded)
                {
                    result.IsIncluded = false;
                    skippedCount++;
                }
                continue;
            }

            foreach (var request in result.Requests.Where(request => request.IsIncluded && !request.IsValid))
            {
                request.IsIncluded = false;
                skippedCount++;
            }
        }

        StatusText.Text = skippedCount == 0
            ? "スキップが必要な依頼はありません"
            : $"エラー・未対応の依頼を{skippedCount}件スキップしました";
        UpdatePdfGenerationState();
    }

    private void RestoreAllRequests_Click(object sender, RoutedEventArgs e)
    {
        ConfirmResultsCheckBox.IsChecked = false;
        foreach (var result in _recognitionResults)
        {
            result.IsIncluded = true;
            foreach (var request in result.Requests)
            {
                request.IsIncluded = true;
            }
        }

        StatusText.Text = "すべてのレポートとテストを処理対象に戻しました";
        UpdatePdfGenerationState();
    }

    private async void SaveSkippedLogButton_Click(object sender, RoutedEventArgs e)
    {
        var logPath = await SaveSkippedLogAsync(showWhenEmpty: true);
        if (logPath is null)
        {
            return;
        }

        StatusText.Text = $"スキップ一覧を保存しました：{Path.GetFileName(logPath)}";
        MessageBox.Show(this,
            $"スキップ一覧を保存しました。\n\n{logPath}",
            "スキップ一覧を保存しました",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void AttentionFilterButton_Click(object sender, RoutedEventArgs e)
    {
        _showOnlyNeedsAttention = !_showOnlyNeedsAttention;
        _recognitionResultsView.Filter = _showOnlyNeedsAttention
            ? item => item is RecognitionResultItem result && result.NeedsAttention
            : null;
        AttentionFilterButton.Content = _showOnlyNeedsAttention
            ? "すべて表示"
            : "要確認だけ表示";
        _recognitionResultsView.Refresh();
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

        var expectedRequestCount = _recognitionResults
            .Where(result => result.IsIncluded)
            .Sum(result => result.Requests.Count(request => request.IsIncluded));
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
        string? skippedLogPath = null;

        try
        {
            skippedLogPath = await SaveSkippedLogAsync(showWhenEmpty: false);
            for (var index = 0; index < requests.Count; index++)
            {
                StatusText.Text = $"PDFを作成中：{index + 1}/{requests.Count}";
                var result = await ExcelPdfGenerationService.GenerateAsync(requests[index]);
                generatedFiles.Add(result.ProblemPdfPath);
                generatedFiles.Add(result.AnswerPdfPath);
            }

            StatusText.Text = $"PDF作成完了：{generatedFiles.Count}ファイル";
            MessageBox.Show(this,
                $"問題PDFと解答PDFを作成しました。\n\n保存先：{OutputFolderTextBox.Text}" +
                (skippedLogPath is null ? string.Empty : $"\nスキップ一覧：{skippedLogPath}"),
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
        SaveSkippedLogButton.IsEnabled = GetSkippedLogEntries(DateTimeOffset.Now).Count > 0;
        CreatePdfsButton.IsEnabled =
            ConfirmResultsCheckBox.IsChecked == true &&
            _recognitionResults.Count > 0 &&
            _recognitionResults.All(result => result.IsReady) &&
            _recognitionResults.Any(result => result.HasIncludedRequests);
    }

    private void InvalidateRecognitionResults()
    {
        _recognitionResults.Clear();
        ConfirmResultsCheckBox.IsChecked = false;
        UpdateRecognitionResultState();
        UpdatePdfGenerationState();
    }

    private List<ExcelPdfGenerationRequest> BuildPdfGenerationRequests()
    {
        var requests = new List<ExcelPdfGenerationRequest>();
        foreach (var result in _recognitionResults)
        {
            if (!result.IsIncluded)
            {
                continue;
            }

            foreach (var editableRequest in result.Requests.Where(request => request.IsIncluded))
            {
                var testRequest = editableRequest.BuildCandidate();
                if (testRequest is null)
                {
                    continue;
                }

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
                    result.StudentName.Trim(),
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

    private async Task<string?> SaveSkippedLogAsync(bool showWhenEmpty)
    {
        var outputFolder = OutputFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
        {
            if (showWhenEmpty)
            {
                MessageBox.Show(this,
                    "先に出力フォルダーを設定してください。",
                    "出力フォルダーが必要です",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            return null;
        }

        var now = DateTimeOffset.Now;
        var entries = GetSkippedLogEntries(now);
        if (entries.Count == 0)
        {
            if (showWhenEmpty)
            {
                MessageBox.Show(this,
                    "現在スキップされているレポートやテストはありません。",
                    "スキップ項目はありません",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            return null;
        }

        try
        {
            return await SkippedTestLogService.SaveAsync(outputFolder, entries, now);
        }
        catch (Exception exception)
        {
            MessageBox.Show(this,
                $"スキップ一覧を保存できませんでした。\n\n{exception.Message}",
                "ログを保存できませんでした",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return null;
        }
    }

    private List<SkippedTestLogEntry> GetSkippedLogEntries(DateTimeOffset loggedAt)
    {
        var entries = new List<SkippedTestLogEntry>();
        foreach (var result in _recognitionResults)
        {
            if (!result.IsIncluded)
            {
                if (result.Requests.Count == 0)
                {
                    entries.Add(new SkippedTestLogEntry(
                        loggedAt,
                        result.StudentName.Trim(),
                        result.FileName,
                        null,
                        string.Empty,
                        string.Empty,
                        string.Empty,
                        result.HasError ? result.NormalTestSummary : "レポート単位でスキップ"));
                }
                else
                {
                    entries.AddRange(result.Requests.Select(request => CreateSkippedLogEntry(
                        loggedAt,
                        result,
                        request,
                        "レポート単位でスキップ")));
                }
                continue;
            }

            entries.AddRange(result.Requests
                .Where(request => !request.IsIncluded)
                .Select(request => CreateSkippedLogEntry(
                    loggedAt,
                    result,
                    request,
                    request.IsValid ? "手動でスキップ" : request.ValidationMessage)));
        }

        return entries;
    }

    private static SkippedTestLogEntry CreateSkippedLogEntry(
        DateTimeOffset loggedAt,
        RecognitionResultItem result,
        EditableTestRequestItem request,
        string reason) => new(
            loggedAt,
            result.StudentName.Trim(),
            result.FileName,
            request.DisplayNumber,
            request.MaterialName.Trim(),
            $"{request.StartNumber.Trim()}-{request.EndNumber.Trim()}",
            request.QuestionCount.Trim(),
            reason);

    private void AddRecognitionResult(RecognitionResultItem result)
    {
        result.Edited += RecognitionResult_Edited;
        _recognitionResults.Add(result);
    }

    private void RecognitionResult_Edited(object? sender, EventArgs e)
    {
        ConfirmResultsCheckBox.IsChecked = false;
        StatusText.Text = "修正内容を原本と照合し、確認チェックを入れ直してください";
        _recognitionResultsView.Refresh();
        UpdatePdfGenerationState();
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
