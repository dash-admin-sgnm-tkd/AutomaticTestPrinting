using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
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
    private bool _isChoosingWorkbook;

    private void ApplyMaterialSuggestion_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: EditableTestRequestItem item } ||
            item.SelectedMaterialSuggestion is not { } selected) return;
        if (MessageBox.Show(this,
            $"教材名を「{item.MaterialName}」から「{selected}」へ変更します。\n原本の教材・版と一致していますか？\n範囲・問題数は変更しません。",
            "教材名の確認", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            item.MaterialName = selected;
    }

    private async void ChooseWorkbook_Click(object sender, RoutedEventArgs e)
    {
        if (_isChoosingWorkbook || _isAddingMaterialFolders || !ValidateButton.IsEnabled ||
            sender is not FrameworkElement { DataContext: EditableTestRequestItem item } ||
            item.WorkbookProfile is not { } profile || item.WorkbookCandidates.Count == 0)
            return;
        var roots = GetAvailableMaterialFolders();
        var candidates = item.WorkbookCandidates.ToArray();
        var window = new WorkbookChoiceWindow(candidates, profile.DisplayName) { Owner = this };
        if (window.ShowDialog() != true || window.SelectedPath is not { } selectedPath) return;
        _isChoosingWorkbook = true;
        ValidateButton.IsEnabled = false;
        UpdateRecognitionResultState();
        var affected = _recognitionResults.SelectMany(report => report.Requests)
            .Where(request => request.WorkbookProfile?.Id == profile.Id)
            .Select(request => (Item: request, Candidate: request.BuildCandidate())).ToArray();
        StatusText.Text = "使用Excelを保存・確認中（この教材の依頼だけ確認します）";
        try
        {
            var relative = roots.Select(root => WorkbookResolutionService.GetRelativePath(root, selectedPath))
                .FirstOrDefault(path => path is not null) ?? throw new InvalidOperationException("教材フォルダー内のExcelを選択してください。");
            var results = await Task.Run(() =>
            {
                if (!candidates.Contains(selectedPath, StringComparer.OrdinalIgnoreCase))
                    throw new InvalidOperationException("表示中の候補からExcelを選択してください。");
                WorkbookChoiceStore.Save(profile.Id, new WorkbookChoice(relative, WorkbookResolutionService.GetHash(selectedPath)));
                return affected.Where(entry => entry.Candidate is not null)
                    .Select(entry => entry.Candidate!).Distinct()
                    .ToDictionary(candidate => candidate, candidate => ExcelTemplateCatalog.Prepare(candidate, roots, candidates));
            });
            foreach (var entry in affected)
            {
                if (entry.Candidate is not null && entry.Item.BuildCandidate() == entry.Candidate &&
                    results.TryGetValue(entry.Candidate, out var result))
                    entry.Item.ApplyExcelPreparation(result);
            }
            ConfirmResultsCheckBox.IsChecked = false;
            _recognitionResultsView.Refresh();
            StatusText.Text = "使用Excelを保存しました。対象教材の確認結果を更新しました";
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or InvalidOperationException)
        {
            MessageBox.Show(this, $"使用Excelを保存できませんでした：{exception.Message}", "Excelの選択");
        }
        finally
        {
            _isChoosingWorkbook = false;
            ValidateButton.IsEnabled = true;
            UpdateRecognitionResultState();
        }
    }
    private readonly ObservableCollection<ReportFile> _reports = [];
    private readonly ObservableCollection<string> _materialFolders = [];
    private readonly ObservableCollection<RecognitionResultItem> _recognitionResults = [];
    private readonly ICollectionView _recognitionResultsView;
    private readonly JsonSettingsStore _settingsStore;
    private bool _showOnlyNeedsAttention;
    private string? _batchPrintPdfPath;

    public MainWindow()
    {
        InitializeComponent();

        var settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AutomaticTestPrinting",
            "settings.json");

        _settingsStore = new JsonSettingsStore(settingsPath);
        _recognitionResultsView = CollectionViewSource.GetDefaultView(_recognitionResults);
        DataContext = new
        {
            Reports = _reports,
            MaterialFolders = _materialFolders,
            RecognitionResultsView = _recognitionResultsView
        };
        UpdateReportListState();
        UpdateRecognitionResultState();
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        var settings = await _settingsStore.LoadAsync();
        ReportInboxFolderTextBox.Text = settings.ReportInboxFolder ?? string.Empty;
        var savedMaterialFolders = settings.MaterialFolders.Count > 0
            ? settings.MaterialFolders
            : string.IsNullOrWhiteSpace(settings.MaterialFolder)
                ? []
                : [settings.MaterialFolder];
        foreach (var folder in savedMaterialFolders
                     .Where(folder => !string.IsNullOrWhiteSpace(folder))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            _materialFolders.Add(folder);
        }
        OutputFolderTextBox.Text = settings.OutputFolder ?? string.Empty;
        UpdateStatus();
    }

    private void SharedCatalog_Click(object sender, RoutedEventArgs e)
    {
        var window = new SharedMaterialCatalogWindow(GetAvailableMaterialFolders()) { Owner = this };
        window.ShowDialog();
        if (!window.WasSaved) return;
        ConfirmResultsCheckBox.IsChecked = false;
        foreach (var report in _recognitionResults)
            foreach (var request in report.Requests) request.RefreshValidation();
        UpdateRecognitionResultState();
        UpdatePdfGenerationState();
    }

    private void PdfMaterials_Click(object sender, RoutedEventArgs e)
    {
        new PdfMaterialWindow(GetAvailableMaterialFolders(), OutputFolderTextBox.Text) { Owner = this }.ShowDialog();
    }

    private void RegisterPdfMaterial_Click(object sender, RoutedEventArgs e)
    {
        var window = new PdfMaterialRegistrationWindow(GetAvailableMaterialFolders()) { Owner = this };
        window.ShowDialog();
        if (!window.WasSaved) return;
        ConfirmResultsCheckBox.IsChecked = false;
        foreach (var report in _recognitionResults)
            foreach (var request in report.Requests) request.RefreshValidation();
        UpdateRecognitionResultState();
        UpdatePdfGenerationState();
    }

    private async void AddMaterialFolder_Click(object sender, RoutedEventArgs e)
    {
        if (_isAddingMaterialFolders || _isChoosingWorkbook || !ValidateButton.IsEnabled) return;
        var folder = SelectFolder("Excel・PDF教材を保存しているフォルダーを選択してください");
        if (folder is null)
        {
            return;
        }

        var fullPath = Path.GetFullPath(folder);
        if (_materialFolders.Contains(fullPath, StringComparer.OrdinalIgnoreCase))
        {
            MessageBox.Show(this,
                "この教材フォルダーはすでに追加されています。",
                "追加済みです",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        _materialFolders.Add(fullPath);
        InvalidateRecognitionResults();
        UpdateStatus();
        await SaveSettingsAsync();
    }

    private async void RemoveMaterialFolder_Click(object sender, RoutedEventArgs e)
    {
        if (_isAddingMaterialFolders || _isChoosingWorkbook || !ValidateButton.IsEnabled) return;
        if (MaterialFoldersListBox.SelectedItem is not string folder)
        {
            MessageBox.Show(this,
                "一覧から削除する教材フォルダーを選択してください。",
                "フォルダーを選択してください",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        _materialFolders.Remove(folder);
        InvalidateRecognitionResults();
        UpdateStatus();
        await SaveSettingsAsync();
    }

    private bool _isAddingMaterialFolders;

    private void MaterialFolders_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = !_isAddingMaterialFolders && !_isChoosingWorkbook && ValidateButton.IsEnabled &&
                    e.Data.GetDataPresent(DataFormats.FileDrop) && (e.AllowedEffects & DragDropEffects.Copy) != 0
            ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private async void MaterialFolders_Drop(object sender, DragEventArgs e)
    {
        e.Handled = true;
        e.Effects = DragDropEffects.None;
        if (_isAddingMaterialFolders || _isChoosingWorkbook || !ValidateButton.IsEnabled ||
            e.Data.GetData(DataFormats.FileDrop) is not string[] paths) return;
        _isAddingMaterialFolders = true;
        ValidateButton.IsEnabled = false;
        UpdatePdfGenerationState();
        var existing = _materialFolders.ToArray();
        StatusText.Text = "ドロップしたフォルダーを確認しています…";
        try
        {
            var result = await Task.Run(() => MaterialFolderDropService.Plan(paths, existing));
            foreach (var folder in result.AddedFolders) _materialFolders.Add(folder);
            if (result.AddedFolders.Count > 0)
            {
                InvalidateRecognitionResults();
                await SaveSettingsAsync();
            }
            UpdateStatus();
            StatusText.Text = $"教材フォルダーを{result.AddedFolders.Count}件追加しました" +
                (result.DuplicateCount > 0 ? $"／登録済み{result.DuplicateCount}件は追加せず" : "") +
                (result.IgnoredCount > 0 ? $"／ファイル・利用できない項目{result.IgnoredCount}件は対象外" : "");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
            MessageBox.Show(this, $"教材フォルダーの追加・設定保存に失敗しました：{exception.Message}", "フォルダー追加");
            StatusText.Text = "設定を保存できませんでした。フォルダー一覧を確認してください";
        }
        finally
        {
            _isAddingMaterialFolders = false;
            ValidateButton.IsEnabled = true;
            UpdatePdfGenerationState();
        }
    }

    private void RegisterMaterial_Click(object sender, RoutedEventArgs e)
    {
        var materialFolders = GetAvailableMaterialFolders();
        if (materialFolders.Length == 0)
        {
            MessageBox.Show(this,
                "先に教材フォルダーを設定してください。",
                "教材フォルダーが必要です",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var registrationWindow = new BulkMaterialRegistrationWindow(materialFolders)
        {
            Owner = this
        };
        if (registrationWindow.ShowDialog() != true)
        {
            return;
        }

        InvalidateRecognitionResults();
        StatusText.Text = $"教材を{registrationWindow.RegisteredCount}件登録しました";
        MessageBox.Show(this,
            $"教材を{registrationWindow.RegisteredCount}件登録しました。\n" +
            $"{registrationWindow.RegisteredDisplayNames}\n\n" +
            "次回からレポートの読み取り候補として使用されます。",
            "教材をまとめて登録しました",
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
        InvalidateBatchOutputs();
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
        InvalidateBatchOutputs();
        UpdateReportListState();
        UpdateRecognitionResultState();
        UpdateStatus();
    }

    private async void ValidateButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy) return;
        using var busy = BeginBusy("レポートを読み取っています");
        var validation = SetupValidator.ValidateFolders(
            GetAvailableMaterialFolders(),
            OutputFolderTextBox.Text,
            _reports);

        if (!validation.IsValid)
        {
            MessageBox.Show(this, validation.Message, "設定を確認してください",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try { await SaveSettingsAsync(); }
        catch (Exception exception)
        {
            StatusText.Text = "設定を保存できませんでした";
            MessageBox.Show(this, exception.Message, "設定の保存", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        ValidateButton.IsEnabled = false;
        _recognitionResults.Clear();
        ConfirmResultsCheckBox.IsChecked = false;
        UpdateRecognitionResultState();

        try
        {
            var reports = _reports.ToArray();
            var materialFolders = GetAvailableMaterialFolders();
            var completedCount = 0;
            var readingCompleted = false;
            SetBusyProgress("レポートを読み取っています", 0, reports.Length);
            IProgress<string> progress = new Progress<string>(message =>
            {
                if (!_isBusy || readingCompleted) return;
                StatusText.Text = message;
                SetBusyProgress("レポートを読み取っています", completedCount, reports.Length);
            });
            using var concurrencyGate = new SemaphoreSlim(2);
            var recognitionTasks = reports.Select(async report =>
            {
                await concurrencyGate.WaitAsync();
                try
                {
                    var service = new WindowsReportOcrService();
                    var recognized = await service.RecognizeAsync(report.FullPath, progress);
                    return await Task.Run(() => RecognitionResultItem.Success(recognized, materialFolders));
                }
                catch (Exception exception)
                {
                    return RecognitionResultItem.Failure(report.FullPath, exception);
                }
                finally
                {
                    concurrencyGate.Release();
                    var current = Interlocked.Increment(ref completedCount);
                    progress.Report($"読み取り中：{current}/{reports.Length}件完了");
                }
            }).ToArray();

            var recognitionResults = await Task.WhenAll(recognitionTasks);
            readingCompleted = true;
            foreach (var result in recognitionResults)
            {
                AddRecognitionResult(result);
            }

            UpdateRecognitionResultState();
            UpdatePdfGenerationState();

            var failedCount = _recognitionResults.Count(result => result.HasError);
            var errorCount = GetCurrentErrorLogEntries(DateTimeOffset.Now).Count;
            var errorLogPath = await SaveErrorLogAsync(showWhenEmpty: false);
            StatusText.Text = failedCount == 0
                ? errorCount == 0
                    ? $"読み取り完了：{_recognitionResults.Count}件"
                    : $"読み取り完了：{_recognitionResults.Count}件、要確認 {errorCount}件" +
                      (errorLogPath is null ? string.Empty : "（詳細ログ保存済み）")
                : $"読み取り完了：成功 {_recognitionResults.Count - failedCount}件、要確認 {errorCount}件" +
                  (errorLogPath is null ? string.Empty : "（詳細ログ保存済み）");
            if (_recognitionResults.Any(result => result.HasReadingWarning) && !_closeAfterBusy)
                MessageBox.Show(this, "依頼欄があるのに通常テストが0件のレポートがあります。\n依頼なしとは判断せず、「要確認だけ表示」と原本PDFを確認してください。詳細はエラーCSVにも記録しています。",
                    "読み取り漏れの可能性", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception exception)
        {
            StatusText.Text = "レポートの読み取り中にエラーが発生しました";
            MessageBox.Show(this, exception.Message, "レポートの読み取り", MessageBoxButton.OK, MessageBoxImage.Error);
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

    private async void SaveErrorLogButton_Click(object sender, RoutedEventArgs e)
    {
        var logPath = await SaveErrorLogAsync(showWhenEmpty: true);
        if (logPath is null)
        {
            return;
        }

        StatusText.Text = $"エラー一覧を保存しました：{Path.GetFileName(logPath)}";
        MessageBox.Show(this,
            $"エラーの詳細一覧を保存しました。\n\n{logPath}",
            "エラー一覧を保存しました",
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
        if (_isBusy) return;
        using var busy = BeginBusy("テストの作成内容を確認しています");
        StatusText.Text = "教材と依頼内容を照合しています…";
        List<TestGenerationPlan> requests;
        try { requests = await BuildPdfGenerationRequestsAsync(); }
        catch (Exception exception)
        {
            StatusText.Text = "教材の照合中にエラーが発生しました";
            MessageBox.Show(this, exception.Message, "テスト作成の準備", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        if (requests.Count == 0)
        {
            MessageBox.Show(this,
                "作成できるテストがありません。教材フォルダーと読み取り結果を確認してください。",
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
                "未対応または入力エラーのテストがあります。教材の連携表示を確認してください。",
                "PDFを作成できません",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var confirmation = MessageBox.Show(this,
            $"テスト {requests.Count}件の個別PDFと一括印刷用PDFを作成します。\n" +
            "レポートごとに「問題 → 解答」の順で一括PDFにまとめます。\n" +
            "作成後に内容を確認してから印刷してください。続けますか？",
            "一括PDFを作成します",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        CreatePdfsButton.IsEnabled = false;
        ValidateButton.IsEnabled = false;
        ConfirmResultsCheckBox.IsEnabled = false;
        InvalidateBatchOutputs();
        var generatedTests = new List<GeneratedReportTest>();
        var outputFolder = OutputFolderTextBox.Text;
        var generationErrors = new List<ErrorLogEntry>();
        string? skippedLogPath = null;
        string? errorLogPath = null;

        try
        {
            skippedLogPath = await SaveSkippedLogAsync(showWhenEmpty: false);
            for (var index = 0; index < requests.Count; index++)
            {
                StatusText.Text = $"PDFを作成中：{index + 1}/{requests.Count}";
                SetBusyProgress("テストPDFを作成しています", index, requests.Count);
                var activeRequest = requests[index];
                try
                {
                    var result = activeRequest.ExcelRequest is not null
                        ? await ExcelPdfGenerationService.GenerateAsync(activeRequest.ExcelRequest)
                        : await Task.Run(() => PdfMaterialCatalog.Generate(activeRequest, outputFolder));
                    generatedTests.Add(new GeneratedReportTest(
                        activeRequest.SourceReportPath ?? activeRequest.SourceReportFileName, result));
                }
                catch (Exception exception)
                {
                    generationErrors.Add(CreateGenerationErrorLogEntry(
                        DateTimeOffset.Now,
                        activeRequest,
                        exception));
                }
            }

            if (generationErrors.Count > 0)
            {
                errorLogPath = await SaveErrorEntriesAsync(
                    generationErrors,
                    DateTimeOffset.Now,
                    showWhenEmpty: false);
            }

            if (generatedTests.Count == 0)
            {
                throw new InvalidOperationException(
                    "すべてのテストでPDF作成に失敗しました。詳細ログを確認してください。");
            }

            StatusText.Text = $"一括PDFを結合中：{generatedTests.Count}件";
            BusyTitleText.Text = "一括印刷用PDFをまとめています";
            BusyProgressBar.IsIndeterminate = true;
            BusyCountText.Text = $"個別PDFの処理完了：{requests.Count}件（結合・保存中）";
            var batchResult = await Task.Run(() => PdfBatchMergeService.MergeByReport(
                generatedTests,
                outputFolder,
                DateTimeOffset.Now,
                duplex: requests.Any(request => request.PdfQuestions is not null)));
            _batchPrintPdfPath = batchResult.PrintPdfPath;
            BatchPdfActionsPanel.Visibility = Visibility.Visible;

            StatusText.Text = generationErrors.Count == 0
                ? $"一括PDF作成完了：{generatedTests.Count}件"
                : $"一括PDF作成完了：成功 {generatedTests.Count}件、失敗 {generationErrors.Count}件";
            if (!_closeAfterBusy) MessageBox.Show(this,
                "レポートごとに「問題 → 解答」の順で一括PDFを作成しました。\n" +
                "段階突破は問題・解答一体、解答未登録の教材は問題のみです。\n" +
                $"成功：{generatedTests.Count}件" +
                (generationErrors.Count == 0 ? string.Empty : $"\n失敗：{generationErrors.Count}件") +
                $"\n\n保存先：{Path.GetDirectoryName(batchResult.PrintPdfPath)}" +
                (skippedLogPath is null ? string.Empty : $"\nスキップ一覧：{skippedLogPath}") +
                (errorLogPath is null ? string.Empty : $"\nエラー詳細：{errorLogPath}"),
                "一括PDF作成完了",
                MessageBoxButton.OK,
                generationErrors.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }
        catch (Exception exception)
        {
            var displayedErrorLogPath = errorLogPath;
            if (generatedTests.Count > 0)
            {
                displayedErrorLogPath =
                    await SaveGenerationErrorLogAsync(null, exception) ?? errorLogPath;
            }
            StatusText.Text = "一括PDF作成中にエラーが発生しました";
            MessageBox.Show(this,
                exception.Message +
                (displayedErrorLogPath is null
                    ? string.Empty
                    : $"\n\n詳細ログ：{displayedErrorLogPath}"),
                "一括PDFを作成できませんでした",
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

    private void OpenBatchPrintPdf_Click(object sender, RoutedEventArgs e) =>
        OpenBatchPdf(_batchPrintPdfPath, "レポート順の印刷用一括PDF");

    private void OpenBatchPdf(string? path, string displayName)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            MessageBox.Show(this,
                $"{displayName}が見つかりません。もう一度一括PDFを作成してください。",
                "PDFを開けません",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            InvalidateBatchOutputs();
            return;
        }

        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
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
            MaterialFolder = _materialFolders.FirstOrDefault(),
            MaterialFolders = _materialFolders.ToList(),
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
        SaveErrorLogButton.IsEnabled = GetCurrentErrorLogEntries(DateTimeOffset.Now).Count > 0;
        CreatePdfsButton.IsEnabled =
            !_isAddingMaterialFolders &&
            !_isChoosingWorkbook &&
            ConfirmResultsCheckBox.IsChecked == true &&
            _recognitionResults.Count > 0 &&
            _recognitionResults.All(result => result.IsReady) &&
            _recognitionResults.Any(result => result.HasIncludedRequests);
    }

    private void InvalidateRecognitionResults()
    {
        _recognitionResults.Clear();
        ConfirmResultsCheckBox.IsChecked = false;
        InvalidateBatchOutputs();
        UpdateRecognitionResultState();
        UpdatePdfGenerationState();
    }

    private void InvalidateBatchOutputs()
    {
        _batchPrintPdfPath = null;
        BatchPdfActionsPanel.Visibility = Visibility.Collapsed;
    }

    private async Task<List<TestGenerationPlan>> BuildPdfGenerationRequestsAsync()
    {
        // Snapshot editable UI state before preparing files on a background thread.
        var entries = _recognitionResults.Where(result => result.IsIncluded)
            .SelectMany(result => result.Requests.Where(request => request.IsIncluded)
                .Select(request => (Candidate: request.BuildCandidate(), StudentName: result.StudentName.Trim(),
                    result.FileName, result.SourcePath, request.DisplayNumber))).ToArray();
        var materialFolders = GetAvailableMaterialFolders();
        var outputFolder = OutputFolderTextBox.Text;
        return await Task.Run(() =>
        {
        var requests = new List<TestGenerationPlan>();
        foreach (var entry in entries)
        {
                var testRequest = entry.Candidate;
                if (testRequest is null)
                {
                    continue;
                }

                var pdf = PdfMaterialCatalog.Prepare(testRequest, materialFolders);
                if (pdf.IsSupported)
                {
                    if (pdf.IsValid && pdf.Questions is not null)
                        requests.Add(new TestGenerationPlan(entry.StudentName, testRequest.MaterialName,
                            testRequest.Range, testRequest.QuestionCount, entry.FileName, entry.DisplayNumber,
                            PdfQuestions: pdf.Questions, SourceReportPath: entry.SourcePath));
                    continue;
                }
                var preparation = ExcelTemplateCatalog.Prepare(
                    testRequest,
                    materialFolders);
                if (!preparation.IsValid ||
                    preparation.Profile is null ||
                    preparation.WorkbookPath is null ||
                    preparation.StartNumber is null ||
                    preparation.EndNumber is null)
                {
                    continue;
                }

                var excelRequest = new ExcelPdfGenerationRequest(
                    entry.StudentName,
                    testRequest.MaterialName,
                    testRequest.QuestionCount,
                    preparation.StartNumber.Value,
                    preparation.EndNumber.Value,
                    preparation.WorkbookPath,
                    outputFolder,
                    preparation.Profile,
                    entry.FileName,
                    entry.DisplayNumber);
                requests.Add(new TestGenerationPlan(entry.StudentName, testRequest.MaterialName,
                    testRequest.Range, testRequest.QuestionCount, entry.FileName, entry.DisplayNumber, excelRequest,
                    SourceReportPath: entry.SourcePath));
        }

        return requests;
        });
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

    private async Task<string?> SaveErrorLogAsync(bool showWhenEmpty)
    {
        var now = DateTimeOffset.Now;
        return await SaveErrorEntriesAsync(
            GetCurrentErrorLogEntries(now),
            now,
            showWhenEmpty);
    }

    private async Task<string?> SaveGenerationErrorLogAsync(
        ExcelPdfGenerationRequest? request,
        Exception exception)
    {
        var now = DateTimeOffset.Now;
        var entry = CreateExceptionErrorLogEntry(
            now,
            "PDF作成",
            request?.StudentName ?? string.Empty,
            request?.SourceReportFileName ?? string.Empty,
            request?.TestNumber,
            request?.MaterialName ?? string.Empty,
            request is null ? string.Empty : $"{request.StartNumber}-{request.EndNumber}",
            request?.QuestionCount.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            exception);
        return await SaveErrorEntriesAsync([entry], now, showWhenEmpty: false);
    }

    private static ErrorLogEntry CreateGenerationErrorLogEntry(
        DateTimeOffset loggedAt,
        TestGenerationPlan request,
        Exception exception) => CreateExceptionErrorLogEntry(
            loggedAt,
            "PDF作成",
            request.StudentName,
            request.SourceReportFileName,
            request.TestNumber,
            request.MaterialName,
            request.Range,
            request.QuestionCount.ToString(CultureInfo.InvariantCulture),
            exception);

    private async Task<string?> SaveErrorEntriesAsync(
        List<ErrorLogEntry> entries,
        DateTimeOffset createdAt,
        bool showWhenEmpty)
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

        if (entries.Count == 0)
        {
            if (showWhenEmpty)
            {
                MessageBox.Show(this,
                    "現在記録するエラーはありません。",
                    "エラーはありません",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            return null;
        }

        try
        {
            return await ErrorLogService.SaveAsync(outputFolder, entries, createdAt);
        }
        catch (Exception exception)
        {
            MessageBox.Show(this,
                $"エラー一覧を保存できませんでした。\n\n{exception.Message}",
                "ログを保存できませんでした",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return null;
        }
    }

    private List<ErrorLogEntry> GetCurrentErrorLogEntries(DateTimeOffset loggedAt)
    {
        var entries = new List<ErrorLogEntry>();
        foreach (var result in _recognitionResults)
        {
            if (result.HasReadingWarning)
                entries.Add(new ErrorLogEntry(loggedAt, "レポート読み取り", result.StudentName.Trim(), result.FileName,
                    null, string.Empty, string.Empty, string.Empty, "RecognitionWarning", string.Empty,
                    result.NormalTestReadWarning, "通常テスト依頼欄は検出されましたが、教材・範囲・問題数の組み合わせを抽出できませんでした。"));
            if (result.HasError)
            {
                entries.Add(new ErrorLogEntry(
                    loggedAt,
                    "レポート読み取り",
                    result.StudentName.Trim(),
                    result.FileName,
                    null,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    result.ErrorType,
                    result.ErrorCode,
                    result.NormalTestSummary,
                    result.ErrorDetails));
            }

            entries.AddRange(result.Requests
                .Where(request => !request.IsValid)
                .Select(request => new ErrorLogEntry(
                    loggedAt,
                    "教材・入力内容の事前検査",
                    result.StudentName.Trim(),
                    result.FileName,
                    request.DisplayNumber,
                    request.MaterialName.Trim(),
                    $"{request.StartNumber.Trim()}-{request.EndNumber.Trim()}",
                    request.QuestionCount.Trim(),
                    "ValidationError",
                    string.Empty,
                    request.ValidationMessage,
                    $"教材名={request.MaterialName}; 開始番号={request.StartNumber}; " +
                    $"終了番号={request.EndNumber}; 問題数={request.QuestionCount}")));
        }

        return entries;
    }

    private static ErrorLogEntry CreateExceptionErrorLogEntry(
        DateTimeOffset loggedAt,
        string stage,
        string studentName,
        string reportFileName,
        int? testNumber,
        string materialName,
        string range,
        string questionCount,
        Exception exception) => new(
            loggedAt,
            stage,
            studentName,
            reportFileName,
            testNumber,
            materialName,
            range,
            questionCount,
            exception.GetType().FullName ?? exception.GetType().Name,
            $"0x{exception.HResult:X8}",
            exception.Message,
            $"使用Excelファイル={exception.Data["SourceWorkbookPath"] ?? "未取得"}\n{exception}");

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
        InvalidateBatchOutputs();
        StatusText.Text = "修正内容を原本と照合し、確認チェックを入れ直してください";
        _recognitionResultsView.Refresh();
        UpdatePdfGenerationState();
    }

    private void UpdateStatus()
    {
        var foldersSelected = GetAvailableMaterialFolders().Length > 0
            && !string.IsNullOrWhiteSpace(OutputFolderTextBox.Text);

        StatusText.Text = foldersSelected && _reports.Count > 0
            ? $"準備確認できます：レポート {_reports.Count}件"
            : "設定を確認してください";
    }

    private static string? EmptyToNull(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    private string[] GetAvailableMaterialFolders() =>
        _materialFolders
            .Where(Directory.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
}
