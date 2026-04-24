using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace RoomManager.Controls;

/// <summary>
/// 虚拟滚动列表控件
/// </summary>
public class VirtualScrollListBox : ListBox
{
    private ScrollViewer? _scrollViewer;
    private bool _isLoading = false;
    private bool _isDisposed = false;

    /// <summary>
    /// 加载更多事件
    /// </summary>
    public event EventHandler? LoadMoreRequested;

    /// <summary>
    /// 是否启用虚拟滚动
    /// </summary>
    public bool IsVirtualScrollEnabled
    {
        get => (bool)GetValue(IsVirtualScrollEnabledProperty);
        set => SetValue(IsVirtualScrollEnabledProperty, value);
    }

    public static readonly DependencyProperty IsVirtualScrollEnabledProperty =
        DependencyProperty.Register(nameof(IsVirtualScrollEnabled), typeof(bool), 
            typeof(VirtualScrollListBox), new PropertyMetadata(true));

    /// <summary>
    /// 加载阈值（距离底部多少像素触发加载）
    /// </summary>
    public double LoadThreshold
    {
        get => (double)GetValue(LoadThresholdProperty);
        set => SetValue(LoadThresholdProperty, value);
    }

    public static readonly DependencyProperty LoadThresholdProperty =
        DependencyProperty.Register(nameof(LoadThreshold), typeof(double), 
            typeof(VirtualScrollListBox), new PropertyMetadata(200.0));

    public VirtualScrollListBox()
    {
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        // 查找 ScrollViewer
        _scrollViewer = GetTemplateChild("ScrollViewer") as ScrollViewer;
        
        if (_scrollViewer == null)
        {
            // 尝试从模板中查找
            ApplyTemplate();
            _scrollViewer = GetTemplateChild("ScrollViewer") as ScrollViewer;
        }

        if (_scrollViewer != null)
        {
            _scrollViewer.ScrollChanged += OnScrollChanged;
        }
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        _isDisposed = true;
        
        if (_scrollViewer != null)
        {
            _scrollViewer.ScrollChanged -= OnScrollChanged;
            _scrollViewer = null;
        }
        
        Loaded -= OnLoaded;
        Unloaded -= OnUnloaded;
    }

    private void OnScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (!IsVirtualScrollEnabled || _isLoading || _isDisposed) return;

        // 检查是否接近底部
        if (_scrollViewer == null) return;

        var verticalOffset = _scrollViewer.VerticalOffset;
        var extentHeight = _scrollViewer.ExtentHeight;
        var viewportHeight = _scrollViewer.ViewportHeight;

        // 如果距离底部小于阈值，触发加载
        if (extentHeight - verticalOffset - viewportHeight < LoadThreshold)
        {
            _isLoading = true;
            LoadMoreRequested?.Invoke(this, EventArgs.Empty);

            // 延迟重置加载状态（使用 Dispatcher 而不是 Task.Run）
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!_isDisposed)
                {
                    _isLoading = false;
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }
    }

    /// <summary>
    /// 重置加载状态
    /// </summary>
    public void ResetLoadingState()
    {
        _isLoading = false;
    }

    /// <summary>
    /// 滚动到顶部
    /// </summary>
    public void ScrollToTop()
    {
        if (_scrollViewer != null)
        {
            _scrollViewer.ScrollToTop();
        }
    }

    /// <summary>
    /// 滚动到底部
    /// </summary>
    public void ScrollToBottom()
    {
        if (_scrollViewer != null)
        {
            _scrollViewer.ScrollToBottom();
        }
    }
}
