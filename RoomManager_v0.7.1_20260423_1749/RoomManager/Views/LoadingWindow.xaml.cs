using System.Windows;
using RoomManager.Services;

namespace RoomManager.Views;

/// <summary>
/// 加载进度窗口
/// </summary>
public partial class LoadingWindow : Window
{
    private readonly AsyncRoomLoader _loader;

    public LoadingWindow(AsyncRoomLoader loader)
    {
        InitializeComponent();
        _loader = loader;

        _loader.ProgressChanged += OnProgressChanged;
        _loader.LoadCompleted += OnLoadCompleted;
        
        // 窗口关闭时取消订阅
        Closed += OnWindowClosed;
    }

    private void OnProgressChanged(object? sender, LoadProgressEventArgs e)
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = e.Percentage;
                ProgressText.Text = $"{e.LoadedCount} / {e.TotalCount} ({e.Percentage:F0}%)";
            });
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"OnProgressChanged 错误: {ex.Message}");
        }
    }

    private void OnLoadCompleted(object? sender, LoadCompleteEventArgs e)
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                if (e.IsCancelled)
                {
                    DialogResult = false;
                }
                else
                {
                    DialogResult = true;
                }

                Close();
            });
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"OnLoadCompleted 错误: {ex.Message}");
        }
    }

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        try
        {
            _loader.ProgressChanged -= OnProgressChanged;
            _loader.LoadCompleted -= OnLoadCompleted;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"OnWindowClosed 错误: {ex.Message}");
        }
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        _loader.Cancel();
        DialogResult = false;
        Close();
    }
}
