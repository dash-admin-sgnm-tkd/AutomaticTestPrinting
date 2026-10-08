using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;
using Microsoft.Win32;

namespace AutomaticTestPrinting.App;

public partial class PdfMaterialRegistrationWindow : Window
{
    private readonly IReadOnlyCollection<string> _folders;
    private readonly ObservableCollection<MappingRow> _rows = [];
    private List<PdfMaterialProfile> _profiles = [];
    private string? _editingId;
    private bool _saving;
    private bool _isStageTest;
    public bool WasSaved { get; private set; }

    public PdfMaterialRegistrationWindow(IReadOnlyCollection<string> folders, PdfMaterialProfile? initialProfile = null)
    {
        InitializeComponent();
        _folders = folders;
        MappingGrid.ItemsSource = _rows;
        try
        {
            _profiles = PdfMaterialProfileStore.Load().ToList();
            ProfilesComboBox.ItemsSource = _profiles;
            if (initialProfile is not null)
            {
                var existing = _profiles.FirstOrDefault(p => p.Id == initialProfile.Id ||
                    p.Aliases.Append(p.DisplayName).Any(a => PdfMaterialProfileStore.NormalizeName(a) == PdfMaterialProfileStore.NormalizeName(initialProfile.DisplayName)));
                if (existing is not null) ProfilesComboBox.SelectedItem = existing;
                else FillProfile(initialProfile);
            }
        }
        catch (Exception exception)
        {
            StatusText.Text = $"設定を読み込めません（上書きを防ぐため保存できません）：{exception.Message}";
            SaveButton.IsEnabled = false;
        }
    }

    private void Profile_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (ProfilesComboBox.SelectedItem is not PdfMaterialProfile profile) return;
        FillProfile(profile);
    }

    private void FillProfile(PdfMaterialProfile profile)
    {
        _editingId = profile.Id;
        _isStageTest = profile.IsStageTest;
        NameTextBox.Text = profile.DisplayName;
        AliasesTextBox.Text = string.Join(',', profile.Aliases);
        _rows.Clear();
        foreach (var question in profile.Questions)
            _rows.Add(new MappingRow(question.Number, question.FileName, question.Pages, question.AnswerFileName, question.AnswerPages, question.ExpectedPageCount));
    }

    private void New_Click(object sender, RoutedEventArgs e)
    {
        ProfilesComboBox.SelectedIndex = -1;
        _editingId = null;
        _isStageTest = false;
        NameTextBox.Clear();
        AliasesTextBox.Clear();
        _rows.Clear();
    }

    private OpenFileDialog PdfDialog(bool multiple) => new()
    {
        Title = "教材フォルダー内のPDFを選択", Filter = "PDF (*.pdf)|*.pdf", Multiselect = multiple,
        InitialDirectory = _folders.FirstOrDefault(Directory.Exists), CheckFileExists = true
    };

    private void Add_Click(object sender, RoutedEventArgs e)
    {
        var dialog = PdfDialog(true);
        if (dialog.ShowDialog(this) != true) return;
        foreach (var file in dialog.FileNames.OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
            _rows.Add(new MappingRow(NextNumber(), Path.GetFileName(file), "", "", ""));
        if (string.IsNullOrWhiteSpace(NameTextBox.Text)) NameTextBox.Text = Path.GetFileNameWithoutExtension(dialog.FileName);
    }

    private int NextNumber() => _rows.Count == 0 ? 1 : checked(_rows.Max(row => row.Number) + 1);

    private void Duplicate_Click(object sender, RoutedEventArgs e)
    {
        if (MappingGrid.SelectedItem is MappingRow row)
            _rows.Add(new MappingRow(NextNumber(), row.FileName, "", row.AnswerFileName, ""));
    }

    private void Answer_Click(object sender, RoutedEventArgs e)
    {
        if (MappingGrid.SelectedItem is not MappingRow row) return;
        var dialog = PdfDialog(false);
        if (dialog.ShowDialog(this) != true) return;
        row.AnswerFileName = Path.GetFileName(dialog.FileName);
        MappingGrid.Items.Refresh();
    }

    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        foreach (var row in MappingGrid.SelectedItems.Cast<MappingRow>().ToArray()) _rows.Remove(row);
    }

    private async void Save_Click(object sender, RoutedEventArgs e)
    {
        if (!MappingGrid.CommitEdit(DataGridEditingUnit.Cell, true) || !MappingGrid.CommitEdit(DataGridEditingUnit.Row, true))
        {
            MessageBox.Show(this, "入力中のセルを確認してください。問題番号は整数で指定します。", "入力エラー");
            return;
        }
        var profile = new PdfMaterialProfile(_editingId ?? Guid.NewGuid().ToString("N"), NameTextBox.Text.Trim(),
            AliasesTextBox.Text.Split([',', '，'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries),
            _rows.Select(row => new PdfQuestionMapping(row.Number, row.FileName.Trim(), row.Pages.Trim(), row.AnswerFileName.Trim(), row.AnswerPages.Trim(), row.ExpectedPageCount)).ToArray(), _isStageTest);
        var updated = _profiles.Where(existing => existing.Id != profile.Id).Append(profile).ToArray();
        SaveButton.IsEnabled = false;
        MappingGrid.IsEnabled = false;
        EditorGrid.IsEnabled = false;
        _saving = true;
        try
        {
            PdfMaterialProfileStore.Validate(updated);
            StatusText.Text = "教材ファイルとページを確認中…";
            await Task.Run(() =>
            {
                var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var question in profile.Questions)
                {
                    var prepared = PdfMaterialCatalog.Prepare(new NormalTestRequestCandidate(profile.DisplayName,
                        $"{question.Number}-{question.Number}", 1, 1, true), _folders, [profile]);
                    if (!prepared.IsValid || prepared.Questions is null) throw new InvalidOperationException(prepared.Message);
                    var resolved = prepared.Questions[0];
                    void Check(string path, string pages)
                    {
                        if (!counts.TryGetValue(path, out var count)) counts[path] = count = PdfMaterialGenerationService.GetPageCount(path);
                        PdfMaterialGenerationService.ParsePages(pages, count);
                    }
                    try
                    {
                        Check(resolved.Path, resolved.Pages);
                        if (resolved.AnswerPath is not null) Check(resolved.AnswerPath, resolved.AnswerPages);
                    }
                    catch (Exception exception)
                    {
                        throw new InvalidOperationException($"問題番号 {question.Number}／{question.FileName}：{exception.Message}", exception);
                    }
                }
                PdfMaterialProfileStore.Save(updated);
            });
            WasSaved = true;
            _saving = false;
            MessageBox.Show(this, "PDF教材を保存しました。読み取り済みの依頼も再検査します。", "登録完了");
            DialogResult = true;
        }
        catch (Exception exception)
        {
            StatusText.Text = exception.Message;
            MessageBox.Show(this, exception.Message, "登録内容を確認してください", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            SaveButton.IsEnabled = true;
            MappingGrid.IsEnabled = true;
            EditorGrid.IsEnabled = true;
            _saving = false;
        }
    }

    private void Window_Closing(object? sender, CancelEventArgs e) => e.Cancel = _saving;

    private sealed class MappingRow(int number, string fileName, string pages, string answerFileName, string answerPages, int? expectedPageCount = null)
    {
        public int Number { get; set; } = number;
        public string FileName { get; set; } = fileName;
        public string Pages { get; set; } = pages;
        public string AnswerFileName { get; set; } = answerFileName;
        public string AnswerPages { get; set; } = answerPages;
        public int? ExpectedPageCount { get; set; } = expectedPageCount;
    }
}
