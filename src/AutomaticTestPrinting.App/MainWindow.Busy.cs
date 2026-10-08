using System.ComponentModel;
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace AutomaticTestPrinting.App;

public partial class MainWindow
{
    private bool _isBusy;
    private bool _closeAfterBusy;
    private bool _allowClose;
    private bool _savingOnClose;

    private BusyScope BeginBusy(string title)
    {
        if (_isBusy) throw new InvalidOperationException("別の処理が実行中です。");
        _isBusy = true;
        BusyTitleText.Text = title;
        BusyCountText.Text = "処理を開始しています…";
        BusyProgressBar.IsIndeterminate = true;
        MainContentGrid.IsEnabled = false;
        BusyOverlay.Visibility = Visibility.Visible;
        BusySpinnerRotation.BeginAnimation(System.Windows.Media.RotateTransform.AngleProperty,
            new DoubleAnimation(0, 360, TimeSpan.FromSeconds(1)) { RepeatBehavior = RepeatBehavior.Forever });
        return new BusyScope(this);
    }

    private void SetBusyProgress(string title, int completed, int total)
    {
        if (!_isBusy) return;
        BusyTitleText.Text = title;
        BusyCountText.Text = $"{completed} / {total} 件完了";
        BusyCloseNotice.Visibility = _closeAfterBusy ? Visibility.Visible : Visibility.Collapsed;
        BusyProgressBar.IsIndeterminate = total <= 0;
        if (total > 0) BusyProgressBar.Value = completed * 100.0 / total;
    }

    private void EndBusy()
    {
        BusySpinnerRotation.BeginAnimation(System.Windows.Media.RotateTransform.AngleProperty, null);
        BusyOverlay.Visibility = Visibility.Collapsed;
        MainContentGrid.IsEnabled = true;
        _isBusy = false;
        if (_closeAfterBusy) Dispatcher.BeginInvoke(new Action(Close), DispatcherPriority.Normal);
    }

    private sealed class BusyScope(MainWindow owner) : IDisposable
    {
        public void Dispose() => owner.EndBusy();
    }

    private async void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_allowClose) return;
        e.Cancel = true;
        if (_isBusy)
        {
            if (!_closeAfterBusy && MessageBox.Show(this,
                $"{BusyTitleText.Text}。途中で閉じると処理が中断されるため、今は終了しません。\n\n処理が完了したら自動で閉じますか？\n「いいえ」を選ぶと、そのまま待機します。",
                "処理中です", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                _closeAfterBusy = true;
                BusyCloseNotice.Visibility = Visibility.Visible;
            }
            return;
        }
        if (_savingOnClose) return;
        _savingOnClose = true;
        try
        {
            await SaveSettingsAsync();
            _allowClose = true;
            _ = Dispatcher.BeginInvoke(new Action(Close), DispatcherPriority.Normal);
        }
        catch (Exception exception)
        {
            _closeAfterBusy = false;
            MessageBox.Show(this, $"設定を保存できなかったため、終了していません。\n{exception.Message}",
                "終了時の設定保存", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally { _savingOnClose = false; }
    }
}
