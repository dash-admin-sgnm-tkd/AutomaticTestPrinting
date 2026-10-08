using System.Windows;
using System.Windows.Controls;

namespace AutomaticTestPrinting.App;

public sealed class WorkbookChoiceWindow : Window
{
    public string? SelectedPath { get; private set; }

    public WorkbookChoiceWindow(IReadOnlyList<string> paths, string materialName)
    {
        Title = "使用Excelの選択：" + materialName;
        Width = 850; Height = 460; MinWidth = 600; MinHeight = 300;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var panel = new DockPanel { Margin = new Thickness(20) };
        var description = new TextBlock
        {
            Text = "使用するExcelを1つ選んでください。保存後は同じ教材すべてに適用し、次回の起動でも引き継ぎます。\nファイルは開かず、マクロ・印刷も実行しません。移動や内容更新時は再確認します。",
            TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 16)
        };
        DockPanel.SetDock(description, Dock.Top); panel.Children.Add(description);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
        var save = new Button { Content = "このExcelを使う（保存）", Padding = new Thickness(16, 8, 16, 8), IsEnabled = false };
        var cancel = new Button { Content = "キャンセル", Padding = new Thickness(16, 8, 16, 8), Margin = new Thickness(12, 0, 0, 0), IsCancel = true };
        buttons.Children.Add(save); buttons.Children.Add(cancel);
        DockPanel.SetDock(buttons, Dock.Bottom); panel.Children.Add(buttons);
        var list = new ListBox { HorizontalContentAlignment = HorizontalAlignment.Stretch };
        foreach (var path in paths)
            list.Items.Add(new ListBoxItem { Tag = path, Content = new TextBlock { Text = path, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(6) } });
        list.SelectionChanged += (_, _) => save.IsEnabled = list.SelectedItem is ListBoxItem;
        save.Click += (_, _) => { SelectedPath = (list.SelectedItem as ListBoxItem)?.Tag as string; DialogResult = true; };
        panel.Children.Add(list); Content = panel;
    }
}
