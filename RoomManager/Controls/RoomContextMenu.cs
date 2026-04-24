using System.Linq;
using System.Windows;
using System.Windows.Controls;
using RoomManager.Models;

namespace RoomManager.Controls;

/// <summary>
/// 房间列表右键菜单
/// </summary>
public class RoomContextMenu : ContextMenu
{
    public static readonly DependencyProperty SelectedRoomProperty =
        DependencyProperty.Register(nameof(SelectedRoom), typeof(RoomData), 
            typeof(RoomContextMenu), new PropertyMetadata(null));

    public static readonly DependencyProperty SelectedRoomsProperty =
        DependencyProperty.Register(nameof(SelectedRooms), typeof(IEnumerable<RoomData>), 
            typeof(RoomContextMenu), new PropertyMetadata(null));

    public RoomData? SelectedRoom
    {
        get => (RoomData?)GetValue(SelectedRoomProperty);
        set => SetValue(SelectedRoomProperty, value);
    }

    public IEnumerable<RoomData>? SelectedRooms
    {
        get => (IEnumerable<RoomData>?)GetValue(SelectedRoomsProperty);
        set => SetValue(SelectedRoomsProperty, value);
    }

    /// <summary>
    /// 编辑房间事件
    /// </summary>
    public event EventHandler<RoomEventArgs>? EditRoomRequested;

    /// <summary>
    /// 删除房间事件
    /// </summary>
    public event EventHandler<RoomEventArgs>? DeleteRoomRequested;

    /// <summary>
    /// 定位房间事件
    /// </summary>
    public event EventHandler<RoomEventArgs>? LocateRoomRequested;

    /// <summary>
    /// 复制房间信息事件
    /// </summary>
    public event EventHandler<RoomEventArgs>? CopyRoomInfoRequested;

    /// <summary>
    /// 批量修改事件
    /// </summary>
    public event EventHandler<RoomsEventArgs>? BatchModifyRequested;

    /// <summary>
    /// 导出选中房间事件
    /// </summary>
    public event EventHandler<RoomsEventArgs>? ExportSelectedRequested;

    public RoomContextMenu()
    {
        CreateMenuItems();
    }

    private void CreateMenuItems()
    {
        // 编辑
        var editItem = new MenuItem
        {
            Header = "编辑房间",
            InputGestureText = "Enter"
        };
        editItem.Click += (s, e) => OnEditRoom();
        Items.Add(editItem);

        Items.Add(new Separator());

        // 定位
        var locateItem = new MenuItem
        {
            Header = "在模型中定位",
            InputGestureText = "Ctrl+L"
        };
        locateItem.Click += (s, e) => OnLocateRoom();
        Items.Add(locateItem);

        // 复制
        var copyItem = new MenuItem
        {
            Header = "复制房间信息",
            InputGestureText = "Ctrl+C"
        };
        copyItem.Click += (s, e) => OnCopyRoomInfo();
        Items.Add(copyItem);

        Items.Add(new Separator());

        // 批量操作
        var batchItem = new MenuItem
        {
            Header = "批量修改"
        };
        
        var batchRenameItem = new MenuItem { Header = "批量重命名" };
        batchRenameItem.Click += (s, e) => OnBatchModify("rename");
        batchItem.Items.Add(batchRenameItem);

        var batchNumberItem = new MenuItem { Header = "批量编号" };
        batchNumberItem.Click += (s, e) => OnBatchModify("number");
        batchItem.Items.Add(batchNumberItem);

        var batchParamItem = new MenuItem { Header = "批量设置参数" };
        batchParamItem.Click += (s, e) => OnBatchModify("parameter");
        batchItem.Items.Add(batchParamItem);

        Items.Add(batchItem);

        Items.Add(new Separator());

        // 导出
        var exportItem = new MenuItem
        {
            Header = "导出选中房间"
        };
        exportItem.Click += (s, e) => OnExportSelected();
        Items.Add(exportItem);

        Items.Add(new Separator());

        // 删除
        var deleteItem = new MenuItem
        {
            Header = "删除房间",
            InputGestureText = "Delete",
            Foreground = System.Windows.Media.Brushes.Red
        };
        deleteItem.Click += (s, e) => OnDeleteRoom();
        Items.Add(deleteItem);
    }

    private void OnEditRoom()
    {
        if (SelectedRoom != null)
        {
            EditRoomRequested?.Invoke(this, new RoomEventArgs { Room = SelectedRoom });
        }
    }

    private void OnLocateRoom()
    {
        if (SelectedRoom != null)
        {
            LocateRoomRequested?.Invoke(this, new RoomEventArgs { Room = SelectedRoom });
        }
    }

    private void OnCopyRoomInfo()
    {
        if (SelectedRoom != null)
        {
            CopyRoomInfoRequested?.Invoke(this, new RoomEventArgs { Room = SelectedRoom });
        }
    }

    private void OnBatchModify(string operationType)
    {
        var rooms = SelectedRooms?.ToList() ?? new List<RoomData>();
        if (rooms.Count > 0)
        {
            BatchModifyRequested?.Invoke(this, new RoomsEventArgs 
            { 
                Rooms = rooms,
                OperationType = operationType
            });
        }
    }

    private void OnExportSelected()
    {
        var rooms = SelectedRooms?.ToList() ?? new List<RoomData>();
        if (rooms.Count > 0)
        {
            ExportSelectedRequested?.Invoke(this, new RoomsEventArgs { Rooms = rooms });
        }
    }

    private void OnDeleteRoom()
    {
        if (SelectedRoom != null)
        {
            var result = MessageBox.Show(
                $"确定要删除房间 \"{SelectedRoom.Name}\" 吗？\n此操作不可撤销。",
                "确认删除",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                DeleteRoomRequested?.Invoke(this, new RoomEventArgs { Room = SelectedRoom });
            }
        }
    }
}

/// <summary>
/// 房间事件参数
/// </summary>
public class RoomEventArgs : EventArgs
{
    public RoomData Room { get; set; } = null!;
}

/// <summary>
/// 多房间事件参数
/// </summary>
public class RoomsEventArgs : EventArgs
{
    public List<RoomData> Rooms { get; set; } = new();
    public string OperationType { get; set; } = "";
}
