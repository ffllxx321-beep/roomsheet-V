using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RoomManager.Controls;
using RoomManager.Models;
using RoomManager.Services;
using RoomManager.ViewModels;

namespace RoomManager.Views;

/// <summary>
/// 主窗口交互逻辑
/// </summary>
public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    private Document? _document;
    private View? _activeView;
    
    /// <summary>
    /// 需要定位的元素 ID（窗口关闭后由命令处理）
    /// </summary>
    public long? LocateElementId { get; set; }

    private static readonly string WindowStateDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoomManager");
    private static readonly string WindowStatePath = Path.Combine(WindowStateDir, "window_state.json");

    public MainWindow()
    {
        try
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;
            
            Closing += MainWindow_Closing;
            Closed += MainWindow_Closed;
            Loaded += MainWindow_Loaded;
        }
        catch (Exception ex)
        {
            var fullError = $"MainWindow 构造函数异常:\n类型: {ex.GetType().FullName}\n消息: {ex.Message}\n堆栈: {ex.StackTrace}";
            if (ex.InnerException != null)
                fullError += $"\n\n内部异常: {ex.InnerException.Message}";
            System.IO.File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoomManager", "mainwindow_crash.log"), fullError);
            throw;
        }
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        RestoreWindowState();
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        try
        {
            // 检查是否有未保存的更改
            if (_viewModel.HasUnsavedChanges)
            {
                var result = MessageBox.Show(
                    "有未保存的更改，是否保存？",
                    "确认关闭",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    _viewModel.Save();
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"MainWindow_Closing 错误: {ex.Message}");
        }
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        try
        {
            // 保存窗口状态
            SaveWindowState();

            // 清理 ViewModel 资源
            _viewModel?.Dispose();
            
            // 清理文档引用
            _document = null;
            _activeView = null;
            
            // 解除事件绑定
            Closing -= MainWindow_Closing;
            Closed -= MainWindow_Closed;
            Loaded -= MainWindow_Loaded;
            
            // 强制垃圾回收
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"MainWindow_Closed 错误: {ex.Message}");
        }
    }

    private void SaveWindowState()
    {
        try
        {
            if (!Directory.Exists(WindowStateDir))
                Directory.CreateDirectory(WindowStateDir);

            var state = new
            {
                Left = Left,
                Top = Top,
                Width = ActualWidth,
                Height = ActualHeight,
                WindowState = WindowState.ToString()
            };
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(state);
            File.WriteAllText(WindowStatePath, json);
        }
        catch { }
    }

    private void RestoreWindowState()
    {
        try
        {
            if (!File.Exists(WindowStatePath)) return;

            var json = File.ReadAllText(WindowStatePath);
            dynamic? state = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
            if (state == null) return;

            double left = state.Left;
            double top = state.Top;
            double width = state.Width;
            double height = state.Height;

            // 确保窗口在屏幕范围内
            if (left >= 0 && top >= 0 && width > 400 && height > 300)
            {
                WindowStartupLocation = WindowStartupLocation.Manual;
                Left = left;
                Top = top;
                Width = width;
                Height = height;
            }

            string ws = state.WindowState;
            if (ws == "Maximized")
                WindowState = WindowState.Maximized;
        }
        catch { }
    }

    public void SetDocument(Document document, View activeView)
    {
        _document = document;
        _activeView = activeView;
        
        // 设置 ListView 按楼层分组
        var view = (System.Windows.Data.CollectionViewSource.GetDefaultView(_viewModel.FilteredRooms));
        view.GroupDescriptions.Add(new System.Windows.Data.PropertyGroupDescription("Level"));
        
        LoadRoomsSync(); // 直接同步加载，不用 async
    }

    /// <summary>
    /// 加载房间数据（同步，符合 Revit API 要求）
    /// </summary>
    private void LoadRoomsSync()
    {
        if (_document == null) return;

        try
        {
            System.Diagnostics.Debug.WriteLine("LoadRoomsSync 开始...");
            _viewModel.IsLoading = true;
            _viewModel.StatusMessage = "正在加载房间数据...";

            // 同步加载房间数据（Revit API 只能在主线程访问）
            var roomService = new RevitRoomService(_document);
            var rooms = new List<RoomData>();

            System.Diagnostics.Debug.WriteLine("开始收集房间元素...");
            var collector = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();
            System.Diagnostics.Debug.WriteLine($"收集器创建完成");

            foreach (Room room in collector)
            {
                try
                {
                    var roomData = roomService.ConvertToRoomData(room);
                    rooms.Add(roomData);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"转换房间失败: {ex.Message}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"收集到 {rooms.Count} 个房间");
            _viewModel.LoadRooms(rooms);
            _viewModel.StatusMessage = $"已加载 {rooms.Count} 间房间";
            System.Diagnostics.Debug.WriteLine("LoadRoomsSync 完成");
        }
        catch (Exception ex)
        {
            var fullError = $"LoadRoomsSync 异常:\n类型: {ex.GetType().FullName}\n消息: {ex.Message}\n堆栈: {ex.StackTrace}";
            if (ex.InnerException != null)
            {
                fullError += $"\n\n内部异常:\n类型: {ex.InnerException.GetType().FullName}\n消息: {ex.InnerException.Message}\n堆栈: {ex.InnerException.StackTrace}";
            }
            System.Diagnostics.Debug.WriteLine(fullError);
            System.IO.File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoomManager", "loadrooms_crash.log"), fullError);
            _viewModel.StatusMessage = $"加载失败: {ex.Message}";
            MessageBox.Show($"加载房间失败: {fullError}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _viewModel.IsLoading = false;
        }
    }

    private void LoadRooms()
    {
        if (_document == null) return;

        try
        {
            _viewModel.IsLoading = true;
            _viewModel.StatusMessage = "正在加载房间数据...";

            var rooms = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Select(r => new RoomData
                {
                    ElementId = r.Id.Value,
                    Name = r.Name,
                    Number = r.Number,
                    Level = r.Level?.Name ?? "",
                    Area = r.Area,
                    Volume = r.Volume,
                    Category = RoomCategoryHelper.Classify(r.Name),
                })
                .ToList();

            _viewModel.LoadRooms(rooms);
            _viewModel.StatusMessage = $"已加载 {rooms.Count} 间房间";
        }
        catch (Exception ex)
        {
            _viewModel.StatusMessage = $"加载失败: {ex.Message}";
            MessageBox.Show($"加载房间失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _viewModel.IsLoading = false;
        }
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            Close();
        }
        else if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            OnSaveRoom(sender, e);
            e.Handled = true;
        }
    }

    private void OnImportFromDwg(object sender, RoutedEventArgs e)
    {
        if (_document == null || _activeView == null)
        {
            MessageBox.Show("请先打开一个视图。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        var wizard = new DwgRecognitionWizard(_document, _activeView);
        wizard.Owner = this;
        wizard.ShowDialog();
    }

    private void OnImportFromExcel(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel 文件|*.xlsx;*.xls",
            Title = "选择要导入的 Excel 文件"
        };

        if (dialog.ShowDialog() == true && _document != null)
        {
            try
            {
                var excelService = new ExcelService();
                var updates = excelService.ImportFromExcel(dialog.FileName);

                if (updates.Count == 0)
                {
                    MessageBox.Show("未找到有效数据", "提示");
                    return;
                }

                int updated = 0;
                using var transaction = new Transaction(_document, "从 Excel 导入房间数据");
                transaction.Start();

                foreach (var update in updates)
                {
                    var elementId = new ElementId(update.ElementId);
                    if (_document.GetElement(elementId) is not Room room) continue;

                    if (!string.IsNullOrEmpty(update.Name))
                        room.Name = update.Name;
                    if (!string.IsNullOrEmpty(update.Number))
                    {
                        var numParam = room.get_Parameter(BuiltInParameter.ROOM_NUMBER);
                        if (numParam != null && !numParam.IsReadOnly)
                            numParam.Set(update.Number);
                    }

                    // 回写自定义参数
                    foreach (var (paramName, paramValue) in update.CustomParameters)
                    {
                        foreach (Parameter param in room.Parameters)
                        {
                            if (param.Definition?.Name == paramName && !param.IsReadOnly)
                            {
                                try
                                {
                                    if (param.StorageType == StorageType.String)
                                        param.Set(paramValue?.ToString() ?? "");
                                    else if (param.StorageType == StorageType.Double)
                                        param.SetValueString(paramValue?.ToString() ?? "");
                                    else if (param.StorageType == StorageType.Integer && int.TryParse(paramValue?.ToString(), out var iv))
                                        param.Set(iv);
                                }
                                catch { }
                                break;
                            }
                        }
                    }
                    updated++;
                }

                transaction.Commit();

                // 重新加载房间数据
                LoadRoomsSync();
                MessageBox.Show($"成功导入并更新 {updated} 间房间", "导入成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnExport(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel 文件|*.xlsx",
            Title = "导出到 Excel"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                var excelService = new ExcelService();
                excelService.ExportToExcel(_viewModel.Rooms.ToList(), dialog.FileName);
                MessageBox.Show("导出成功", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnExportPdf(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "PDF 文件|*.pdf",
            Title = "导出到 PDF"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                var pdfService = new PdfReportService();
                pdfService.GenerateRoomReport(_viewModel.Rooms.ToList(), dialog.FileName);
                MessageBox.Show("导出成功", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnValidate(object sender, RoutedEventArgs e)
    {
        var validationService = new ValidationService();
        var report = validationService.ValidateRooms(_viewModel.Rooms);
        
        var window = new ValidationWindow(report);
        window.Owner = this;
        window.ShowDialog();
    }

    private void OnBatchOperation(object sender, RoutedEventArgs e)
    {
        if (_document == null)
        {
            MessageBox.Show("文档未加载。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        // 使用当前选中的房间（如果有），否则使用所有房间
        var selectedRooms = _viewModel.SelectedRoom != null 
            ? new List<RoomData> { _viewModel.SelectedRoom }
            : _viewModel.Rooms.ToList();
        
        var window = new BatchOperationWindow(_document, selectedRooms);
        window.Owner = this;
        window.ShowDialog();
    }

    private void OnSaveAll(object sender, RoutedEventArgs e)
    {
        if (_document == null) return;

        try
        {
            int totalUpdated = 0;
            int roomsUpdated = 0;

            using var transaction = new Transaction(_document, "批量保存房间参数");
            transaction.Start();

            foreach (var roomData in _viewModel.Rooms)
            {
                // 检查是否有修改的参数
                var modifiedParams = roomData.AllParameters.Where(p => p.IsModified).ToList();
                if (modifiedParams.Count == 0) continue;

                var elementId = new ElementId(roomData.ElementId);
                if (_document.GetElement(elementId) is not Room room) continue;

                foreach (var paramInfo in modifiedParams)
                {
                    foreach (Parameter param in room.Parameters)
                    {
                        if (param.Definition?.Name != paramInfo.Name || param.IsReadOnly) continue;
                        try
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.String:
                                    param.Set(paramInfo.DisplayValue ?? "");
                                    totalUpdated++;
                                    break;
                                case StorageType.Integer:
                                    if (int.TryParse(paramInfo.DisplayValue, out var intVal))
                                    { param.Set(intVal); totalUpdated++; }
                                    break;
                                case StorageType.Double:
                                    param.SetValueString(paramInfo.DisplayValue ?? "");
                                    totalUpdated++;
                                    break;
                            }
                            paramInfo.OriginalValue = paramInfo.DisplayValue;
                        }
                        catch { }
                        break;
                    }
                }
                roomsUpdated++;
            }

            transaction.Commit();
            _viewModel.StatusMessage = $"已保存 {roomsUpdated} 间房间的 {totalUpdated} 个参数";
            _viewModel.HasUnsavedChanges = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"批量保存失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnFieldManager(object sender, RoutedEventArgs e)
    {
        var window = new FieldManagerWindow(_viewModel);
        window.Owner = this;
        window.ShowDialog();
    }

    private void OnRoomSelected(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is RoomData room)
        {
            _viewModel.SelectedRoom = room;
        }
    }

    private void OnRoomListSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ListView listView && listView.SelectedItem is RoomData room)
        {
            _viewModel.SelectedRoom = room;
        }
    }

    private void OnSelectAll(object sender, RoutedEventArgs e)
    {
        foreach (var room in _viewModel.FilteredRooms)
            room.IsSelected = true;
        RoomListView.SelectAll();
    }

    private void OnInvertSelection(object sender, RoutedEventArgs e)
    {
        foreach (var room in _viewModel.FilteredRooms)
            room.IsSelected = !room.IsSelected;
    }

    private void OnBatchEdit(object sender, RoutedEventArgs e)
    {
        if (_document == null) return;

        var selectedRooms = _viewModel.FilteredRooms.Where(r => r.IsSelected).ToList();
        if (selectedRooms.Count == 0)
        {
            selectedRooms = _viewModel.Rooms.ToList();
            MessageBox.Show($"未选中房间，将对全部 {selectedRooms.Count} 间房间执行批量操作。", "提示");
        }

        var window = new BatchOperationWindow(_document, selectedRooms);
        window.Owner = this;
        window.ShowDialog();
    }

    private void OnTreeViewContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_viewModel.SelectedRoom == null)
        {
            e.Handled = true; // 没有选中房间时不弹出右键菜单
        }
    }

    private void OnLoadThumbnail(object sender, RoutedEventArgs e)
    {
        if (_document == null || _viewModel.SelectedRoom == null) return;

        try
        {
            var thumbnailService = new ThumbnailService(_document);
            var thumbnail = thumbnailService.GenerateRoomThumbnail(_viewModel.SelectedRoom.ElementId);
            if (thumbnail != null)
            {
                _viewModel.SelectedRoom.Thumbnail = thumbnail;
                // 触发 UI 更新（重新设置 SelectedRoom 让绑定刷新）
                var current = _viewModel.SelectedRoom;
                _viewModel.SelectedRoom = null;
                _viewModel.SelectedRoom = current;
            }
            else
            {
                MessageBox.Show("无法生成缩略图（房间可能没有边界）", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载缩略图失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnLocateRoom(object sender, RoutedEventArgs e)
    {
        LocateToSelectedRoom();
    }

    private void OnRoomDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        LocateToSelectedRoom();
    }

    private void LocateToSelectedRoom()
    {
        if (_viewModel.SelectedRoom == null) return;
        LocateElementId = _viewModel.SelectedRoom.ElementId;
        Close();
    }

    private void OnSaveRoom(object sender, RoutedEventArgs e)
    {
        if (_document == null || _viewModel.SelectedRoom == null) return;

        try
        {
            var roomData = _viewModel.SelectedRoom;
            var elementId = new ElementId(roomData.ElementId);
            var element = _document.GetElement(elementId);

            if (element is Room room)
            {
                using var transaction = new Transaction(_document, "修改房间信息");
                transaction.Start();

                // 遍历所有可编辑参数，写回 Revit
                int updatedCount = 0;
                foreach (var paramInfo in roomData.AllParameters)
                {
                    if (paramInfo.IsReadOnly) continue;
                    
                    // 查找对应的 Revit 参数
                    foreach (Parameter param in room.Parameters)
                    {
                        if (param.Definition?.Name == paramInfo.Name && !param.IsReadOnly)
                        {
                            try
                            {
                                switch (param.StorageType)
                                {
                                    case StorageType.String:
                                        param.Set(paramInfo.DisplayValue ?? "");
                                        updatedCount++;
                                        break;
                                    case StorageType.Integer:
                                        if (int.TryParse(paramInfo.DisplayValue, out var intVal))
                                        {
                                            param.Set(intVal);
                                            updatedCount++;
                                        }
                                        break;
                                    case StorageType.Double:
                                        if (double.TryParse(paramInfo.DisplayValue, out var dblVal))
                                        {
                                            param.SetValueString(paramInfo.DisplayValue);
                                            updatedCount++;
                                        }
                                        break;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"写入参数 {paramInfo.Name} 失败: {ex.Message}");
                            }
                            break;
                        }
                    }
                }

                transaction.Commit();
                _viewModel.HasUnsavedChanges = false;
                _viewModel.StatusMessage = $"已保存 {updatedCount} 个参数";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}