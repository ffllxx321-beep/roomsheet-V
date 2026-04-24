using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System.Linq;
using RoomManager.Models;
using System.Collections.Concurrent;

namespace RoomManager.Services;

/// <summary>
/// 异步房间数据加载服务
/// </summary>
public class AsyncRoomLoader
{
    private readonly Document _document;
    private readonly RevitRoomService _roomService;
    private CancellationTokenSource? _cancellationTokenSource;
    private ConcurrentBag<RoomData> _loadedRooms = new();

    /// <summary>
    /// 加载进度
    /// </summary>
    public int LoadedCount => _loadedRooms.Count;

    /// <summary>
    /// 总数
    /// </summary>
    public int TotalCount { get; private set; }

    /// <summary>
    /// 是否正在加载
    /// </summary>
    public bool IsLoading { get; private set; }

    /// <summary>
    /// 进度变化事件
    /// </summary>
    public event EventHandler<LoadProgressEventArgs>? ProgressChanged;

    /// <summary>
    /// 加载完成事件
    /// </summary>
    public event EventHandler<LoadCompleteEventArgs>? LoadCompleted;

    public AsyncRoomLoader(Document document)
    {
        _document = document;
        _roomService = new RevitRoomService(document);
    }

    /// <summary>
    /// 异步加载所有房间
    /// </summary>
    public async Task<List<RoomData>> LoadAllRoomsAsync(IProgress<LoadProgressEventArgs>? progress = null)
    {
        // 取消之前的加载
        Cancel();

        _cancellationTokenSource = new CancellationTokenSource();
        _loadedRooms = new ConcurrentBag<RoomData>();
        IsLoading = true;

        try
        {
            // 先获取总数
            var allRoomIds = await Task.Run(() =>
            {
                return new FilteredElementCollector(_document)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .ToElementIds()
                    .ToList();
            }, _cancellationTokenSource.Token);

            TotalCount = allRoomIds.Count;

            // 分批加载
            var batchSize = 20;
            var batches = allRoomIds
                .Select((id, index) => new { id, index })
                .GroupBy(x => x.index / batchSize)
                .Select(g => g.Select(x => x.id).ToList())
                .ToList();

            var loadedCount = 0;

            foreach (var batch in batches)
            {
                if (_cancellationTokenSource.Token.IsCancellationRequested)
                    break;

                // 加载一批房间
                var rooms = await Task.Run(() =>
                {
                    var result = new List<RoomData>();
                    foreach (var id in batch)
                    {
                        if (_cancellationTokenSource.Token.IsCancellationRequested)
                            break;

                        var room = _document.GetElement(id) as Room;
                        if (room != null)
                        {
                            var roomData = _roomService.ConvertToRoomData(room);
                            result.Add(roomData);
                            _loadedRooms.Add(roomData);
                        }
                    }
                    return result;
                }, _cancellationTokenSource.Token);

                loadedCount += rooms.Count;

                // 报告进度
                var progressArgs = new LoadProgressEventArgs
                {
                    LoadedCount = loadedCount,
                    TotalCount = TotalCount,
                    Percentage = (double)loadedCount / TotalCount * 100
                };

                progress?.Report(progressArgs);
                ProgressChanged?.Invoke(this, progressArgs);
            }

            var result = _loadedRooms.ToList();

            // 触发完成事件
            LoadCompleted?.Invoke(this, new LoadCompleteEventArgs
            {
                Rooms = result,
                IsCancelled = _cancellationTokenSource.Token.IsCancellationRequested
            });

            return result;
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>
    /// 异步加载指定楼层的房间
    /// </summary>
    public async Task<List<RoomData>> LoadRoomsByLevelAsync(string levelName)
    {
        return await Task.Run(() =>
        {
            var rooms = new List<RoomData>();

            var collector = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            foreach (Room room in collector)
            {
                if (room.Level?.Name == levelName)
                {
                    rooms.Add(_roomService.ConvertToRoomData(room));
                }
            }

            return rooms;
        });
    }

    /// <summary>
    /// 异步搜索房间
    /// </summary>
    public async Task<List<RoomData>> SearchRoomsAsync(string keyword)
    {
        return await Task.Run(() =>
        {
            var rooms = new List<RoomData>();

            var collector = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            foreach (Room room in collector)
            {
                var name = room.Name ?? "";
                var number = room.Number ?? "";

                if (name.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                    number.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    rooms.Add(_roomService.ConvertToRoomData(room));
                }
            }

            return rooms;
        });
    }

    /// <summary>
    /// 取消加载
    /// </summary>
    public void Cancel()
    {
        if (_cancellationTokenSource == null) return;
        
        try
        {
            _cancellationTokenSource.Cancel();
            _cancellationTokenSource.Dispose();
            _cancellationTokenSource = null;
            IsLoading = false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"AsyncRoomLoader.Cancel 错误: {ex.Message}");
        }
    }
}

/// <summary>
/// 加载进度事件参数
/// </summary>
public class LoadProgressEventArgs : EventArgs
{
    public int LoadedCount { get; set; }
    public int TotalCount { get; set; }
    public double Percentage { get; set; }
}

/// <summary>
/// 加载完成事件参数
/// </summary>
public class LoadCompleteEventArgs : EventArgs
{
    public List<RoomData> Rooms { get; set; } = new();
    public bool IsCancelled { get; set; }
}
