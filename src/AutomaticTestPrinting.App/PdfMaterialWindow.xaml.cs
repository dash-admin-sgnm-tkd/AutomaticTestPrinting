using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using Microsoft.Win32;

namespace AutomaticTestPrinting.App;

public partial class PdfMaterialWindow : Window
{
    private readonly ObservableCollection<PdfJobRow> _jobs = [];
    private readonly string? _initialFolder;
    private string _outputFolder;
    private bool _busy;
    private PdfMaterialBatchResult? _result;

    public PdfMaterialWindow(IReadOnlyCollection<string> materialFolders, string outputFolder)
    {
        InitializeComponent();
        _initialFolder = materialFolders.FirstOrDefault(Directory.Exists);
        _outputFolder = outputFolder;
        JobsGrid.ItemsSource = _jobs;
    }

    private async void AddFiles_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "教材のPDFを選択（複数選択できます）",
            Filter = "PDF (*.pdf)|*.pdf",
            InitialDirectory = _initialFolder,
            Multiselect = true,
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }
        var student = StudentTextBox.Text.Trim();
        SetBusy(true);
        try
        {
            foreach (var path in dialog.FileNames.OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
            {
                StatusText.Text = $"PDFを確認中：{Path.GetFileName(path)}";
                try
                {
                    var count = await Task.Run(() => PdfMaterialGenerationService.GetPageCount(path));
                    _jobs.Add(new PdfJobRow(student, path, count));
                }
                catch (Exception exception)
                {
                    throw new InvalidOperationException($"PDFを読み込めません：{path}\n{exception.Message}", exception);
                }
            }
            StatusText.Text = $"{_jobs.Count}件。行をダブルクリックして生徒名・教材名・ページを修正できます。";
        }
        catch (Exception exception)
        {
            await ShowErrorAsync("PDF教材・読み込み", exception);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        foreach (var row in JobsGrid.SelectedItems.Cast<PdfJobRow>().ToArray())
        {
            _jobs.Remove(row);
        }
    }

    private async void Create_Click(object sender, RoutedEventArgs e)
    {
        JobsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        JobsGrid.CommitEdit(DataGridEditingUnit.Row, true);
        if (_jobs.Count == 0)
        {
            MessageBox.Show(this, "教材PDFを追加してください。", "PDFがありません");
            return;
        }
        if (string.IsNullOrWhiteSpace(_outputFolder) || !Directory.Exists(_outputFolder))
        {
            var dialog = new OpenFolderDialog { Title = "印刷用PDFの保存先を選択" };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }
            _outputFolder = dialog.FolderName;
        }
        try
        {
            foreach (var row in _jobs)
            {
                if (string.IsNullOrWhiteSpace(row.StudentName) || string.IsNullOrWhiteSpace(row.MaterialName))
                {
                    throw new InvalidOperationException($"生徒名と教材名を入力してください：{row.FileName}");
                }
                try
                {
                    PdfMaterialGenerationService.ParsePages(row.Pages, row.PageCount);
                }
                catch (ArgumentException exception)
                {
                    throw new InvalidOperationException($"{row.StudentName}／{row.FileName}：{exception.Message}", exception);
                }
            }
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "入力を確認してください", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (MessageBox.Show(this, $"{_jobs.Count}件の印刷用PDFを作成します。\n追加順とページ指定を確認しましたか？",
                "PDF作成の確認", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK)
        {
            return;
        }

        var requests = _jobs.Select(row => new PdfMaterialRequest(row.StudentName.Trim(), row.MaterialName.Trim(),
            row.SourcePath, row.Pages, row.IsAnswer)).ToArray();
        var duplex = DuplexCheckBox.IsChecked == true;
        _result = null;
        SetBusy(true);
        StatusText.Text = "PDFを作成中…（元PDFは変更しません）";
        try
        {
            _result = await Task.Run(() => PdfMaterialGenerationService.Generate(requests, _outputFolder, duplex));
            StatusText.Text = $"作成：{_result.Outputs.Count}件／失敗：{_result.Errors.Count}件\n保存先：{_result.OutputFolder}";
            if (_result.Errors.Count > 0)
            {
                var log = await ErrorLogService.SaveAsync(_result.OutputFolder, _result.Errors, DateTimeOffset.Now);
                MessageBox.Show(this, $"作成できないPDFが{_result.Errors.Count}件ありました。\n成功分のみ一括PDFに含まれます。\n\nエラーログ：{log}",
                    "一部のPDFを作成できませんでした", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception exception)
        {
            await ShowErrorAsync("PDF教材・作成／結合", exception);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task ShowErrorAsync(string stage, Exception exception)
    {
        var message = exception.Message;
        if (Directory.Exists(_outputFolder))
        {
            try
            {
                var entry = new ErrorLogEntry(DateTimeOffset.Now, stage, StudentTextBox.Text, string.Empty, null,
                    string.Empty, string.Empty, string.Empty, exception.GetType().FullName ?? string.Empty,
                    $"0x{exception.HResult:X8}", exception.Message, exception.ToString());
                var log = await ErrorLogService.SaveAsync(_outputFolder, [entry], DateTimeOffset.Now);
                message += $"\n\nエラーログ：{log}";
            }
            catch (Exception logException)
            {
                message += $"\n\nログ保存も失敗しました：{logException.Message}";
            }
        }
        StatusText.Text = "処理に失敗しました。内容を確認してください。";
        MessageBox.Show(this, message, "PDF教材のエラー", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        SelectionPanel.IsEnabled = !busy;
        JobsGrid.IsEnabled = !busy;
        DuplexCheckBox.IsEnabled = !busy;
        CreateButton.IsEnabled = !busy;
        OpenProblemButton.IsEnabled = !busy && _result?.ProblemPdfPath is not null;
        OpenAnswerButton.IsEnabled = !busy && _result?.AnswerPdfPath is not null;
        OpenFolderButton.IsEnabled = !busy && _result is not null;
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        e.Cancel = _busy;
    }

    private void OpenProblem_Click(object sender, RoutedEventArgs e) => OpenPath(_result?.ProblemPdfPath);
    private void OpenAnswer_Click(object sender, RoutedEventArgs e) => OpenPath(_result?.AnswerPdfPath);
    private void OpenFolder_Click(object sender, RoutedEventArgs e) => OpenPath(_result?.OutputFolder);

    private void OpenPath(string? path)
    {
        if (path is null)
        {
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "ファイルを開けませんでした");
        }
    }

    private sealed class PdfJobRow(string studentName, string sourcePath, int pageCount)
    {
        public string StudentName { get; set; } = studentName;
        public string MaterialName { get; set; } = Path.GetFileNameWithoutExtension(sourcePath);
        public string SourcePath { get; } = sourcePath;
        public string FileName => Path.GetFileName(SourcePath);
        public int PageCount { get; } = pageCount;
        public string Pages { get; set; } = string.Empty;
        public bool IsAnswer { get; set; }
    }
}
