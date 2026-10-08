using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.ComponentModel;
using AutomaticTestPrinting.Core.Models;
using AutomaticTestPrinting.Core.Services;

namespace AutomaticTestPrinting.App;

public partial class SharedMaterialCatalogWindow : Window
{
    private readonly IReadOnlyCollection<string> _folders;
    private ICollectionView? _view;
    public bool WasSaved { get; private set; }

    public SharedMaterialCatalogWindow(IReadOnlyCollection<string> folders)
    {
        InitializeComponent();
        _folders = folders;
        try
        {
            var catalog = SharedMaterialCatalog.Load();
            _view = CollectionViewSource.GetDefaultView(catalog.Items.ToArray());
            _view.Filter = item => item is SharedMaterialEntry entry && Matches(entry);
            MaterialsGrid.ItemsSource = _view;
            KindComboBox.ItemsSource = new[] { "すべて", "Excel", "PDF", "段階突破" };
            KindComboBox.SelectedIndex = 0;
            SummaryText.Text = (string.IsNullOrWhiteSpace(catalog.ScopeDescription) ? "" : $"対象：{catalog.ScopeDescription}\n") +
                $"更新日：{catalog.UpdatedAt} ／ " +
                string.Join(" ／ ", catalog.Items.GroupBy(e => e.Kind).Select(g => $"{g.Key} {g.Count()}件")) +
                $" ／ PDF対応表 {catalog.Items.Count(e => e.Kind == "PDF" && e.AutoEnabled)}件・段階突破対応 {catalog.Items.Count(e => e.Kind == "段階突破" && e.AutoEnabled)}件";
        }
        catch (Exception exception)
        {
            SummaryText.Text = $"カタログを読み込めません：{exception.Message}";
        }
    }

    private bool Matches(SharedMaterialEntry entry)
    {
        var kind = KindComboBox.SelectedItem as string;
        if (kind is not null && kind != "すべて" && entry.Kind != kind) return false;
        var query = PdfMaterialProfileStore.NormalizeName(SearchTextBox.Text);
        return query.Length == 0 || PdfMaterialProfileStore.NormalizeName(
            entry.DisplayName + entry.Subject + entry.Edition + entry.FileNameSummary + string.Join(" ", entry.Aliases)).Contains(query, StringComparison.Ordinal);
    }

    private void Filter_Changed(object sender, RoutedEventArgs e) => _view?.Refresh();

    private void Selection_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (MaterialsGrid.SelectedItem is not SharedMaterialEntry entry) { RegisterButton.IsEnabled = false; return; }
        RegisterButton.IsEnabled = string.IsNullOrWhiteSpace(entry.RegistrationReviewReason);
        DetailText.Text = $"{entry.DisplayName}\nファイル：{entry.FileNameSummary}\n" +
            $"出典：{entry.SourceSheet}" + (entry.SourceRow > 0 ? $" {entry.SourceRow}行" : "") + "\n" +
            (string.IsNullOrWhiteSpace(entry.SourceRelativePath) ? "" : $"原本の保存先（指定フォルダーから）：{entry.SourceRelativePath}\n") +
            (string.IsNullOrWhiteSpace(entry.RegistrationReviewReason) ? "" :
                $"登録保留の理由：{entry.RegistrationReviewReason}\n") +
            $"\n{entry.Notes}\n\n" +
            "ファイルは配布物に含まれません。設定した教材フォルダー（Google Drive・OneDrive・PC内）から検索します。";
    }

    private void Register_Click(object sender, RoutedEventArgs e)
    {
        if (MaterialsGrid.SelectedItem is not SharedMaterialEntry entry) return;
        if (!string.IsNullOrWhiteSpace(entry.RegistrationReviewReason))
        {
            MessageBox.Show(this, entry.RegistrationReviewReason, "教材情報を確認してください");
            return;
        }
        if (_folders.Count == 0) { MessageBox.Show(this, "先にメイン画面で教材フォルダーを指定してください。"); return; }
        if (entry.Kind == "Excel")
        {
            var window = new MaterialRegistrationWindow(_folders, entry) { Owner = this };
            if (window.ShowDialog() == true) WasSaved = true;
        }
        else
        {
            var profile = new PdfMaterialProfile(entry.Id, entry.DisplayName, entry.Aliases, entry.Questions, entry.Kind == "段階突破");
            var window = new PdfMaterialRegistrationWindow(_folders, profile) { Owner = this };
            window.ShowDialog();
            WasSaved |= window.WasSaved;
        }
    }

    private void BulkExcel_Click(object sender, RoutedEventArgs e)
    {
        if (_folders.Count == 0) { MessageBox.Show(this, "先にメイン画面で教材フォルダーを指定してください。"); return; }
        if (new BulkMaterialRegistrationWindow(_folders) { Owner = this }.ShowDialog() == true) WasSaved = true;
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
