using RoomManager.Models;
using RoomManager.Utils;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace RoomManager.ViewModels;

/// <summary>
/// 异步房间列表 ViewModel（支持虚拟滚动）
/// </summary>
public class AsyncRoomListViewModel : INotifyPropertyChanged
{
    private readonly int _pageSize = 50;
    private int _currentPage = 0;
    private bool _isLoading = false;
    private bool _hasMoreItems = true;
    private string _searchText = "";
    private int _totalCount;

    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>
    /// 所有房间数据（原始）
    /// </summary>
    private List<RoomData> _allRooms = new();

    /// <summary>
    /// 当前显示的房间（虚拟化）
    /// </summary>
    public ObservableCollection<RoomData> DisplayedRooms { get; } = new();

    /// <summary>
    /// 楼层分组
    /// </summary>
    public ObservableCollection<LevelGroupViewModel> LevelGroups { get; } = new();

    /// <summary>
    /// 是否正在加载
    /// </summary>
    public bool IsLoading
    {
        get => _isLoading;
        set
        {
            _isLoading = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// 是否有更多数据
    /// </summary>
    public bool HasMoreItems
    {
        get => _hasMoreItems;
        set
        {
            _hasMoreItems = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// 搜索文本
    /// </summary>
    public string SearchText
    {
        get => _searchText;
        set
        {
            if (_searchText != value)
            {
                _searchText = value;
                OnPropertyChanged();
                _ = RefreshAsync();
            }
        }
    }

    /// <summary>
    /// 总房间数
    /// </summary>
    public int TotalCount
    {
        get => _totalCount;
        set
        {
            _totalCount = value;
            OnPropertyChanged();
        }
    }

    /// <summary>
    /// 加载更多命令
    /// </summary>
    public ICommand LoadMoreCommand { get; }

    public AsyncRoomListViewModel()
    {
        LoadMoreCommand = new RelayCommand(async () => await LoadMoreAsync());
    }

    /// <summary>
    /// 设置所有房间数据
    /// </summary>
    public void SetRooms(List<RoomData> rooms)
    {
        _allRooms = rooms;
        TotalCount = rooms.Count;
        _ = RefreshAsync();
    }

    /// <summary>
    /// 刷新数据（搜索或初始加载）
    /// </summary>
    public async Task RefreshAsync()
    {
        _currentPage = 0;
        HasMoreItems = true;

        DisplayedRooms.Clear();
        LevelGroups.Clear();

        await LoadMoreAsync();
    }

    /// <summary>
    /// 加载更多数据
    /// </summary>
    public async Task LoadMoreAsync()
    {
        if (IsLoading || !HasMoreItems) return;

        IsLoading = true;

        try
        {
            // 模拟异步加载（实际应用中从数据库或 Revit 加载）
            await Task.Run(() =>
            {
                // 过滤
                var filtered = string.IsNullOrEmpty(SearchText)
                    ? _allRooms
                    : _allRooms.Where(r => 
                        r.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                        r.Number.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                        r.Level.Contains(SearchText, StringComparison.OrdinalIgnoreCase)).ToList();

                // 分页
                var skip = _currentPage * _pageSize;
                var pageData = filtered.Skip(skip).Take(_pageSize).ToList();

                // 更新 UI（需要在主线程）
                System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                {
                    foreach (var room in pageData)
                    {
                        DisplayedRooms.Add(room);
                    }

                    // 更新楼层分组
                    UpdateLevelGroups(pageData);

                    _currentPage++;
                    HasMoreItems = skip + _pageSize < filtered.Count;
                });
            });
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>
    /// 更新楼层分组
    /// </summary>
    private void UpdateLevelGroups(List<RoomData> newRooms)
    {
        // 按楼层分组
        var groups = newRooms
            .GroupBy(r => r.Level)
            .Select(g => new LevelGroupViewModel
            {
                LevelName = g.Key,
                Rooms = new ObservableCollection<RoomData>(g.ToList())
            });

        foreach (var group in groups)
        {
            // 检查是否已存在该楼层
            var existing = LevelGroups.FirstOrDefault(g => g.LevelName == group.LevelName);
            if (existing != null)
            {
                // 合并房间
                foreach (var room in group.Rooms)
                {
                    if (!existing.Rooms.Any(r => r.ElementId == room.ElementId))
                    {
                        existing.Rooms.Add(room);
                    }
                }
            }
            else
            {
                LevelGroups.Add(group);
            }
        }
    }

    /// <summary>
    /// 获取所有房间（用于导出等操作）
    /// </summary>
    public List<RoomData> GetAllRooms()
    {
        return _allRooms;
    }

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

/// <summary>
/// 楼层分组 ViewModel
/// </summary>
public class LevelGroupViewModel
{
    public string LevelName { get; set; } = "";
    public ObservableCollection<RoomData> Rooms { get; set; } = new();
    public int RoomCount => Rooms.Count;
}
