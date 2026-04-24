using RoomManager.Models;
using RoomManager.Services;
using RoomManager.Utils;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace RoomManager.ViewModels;

/// <summary>
/// 主窗口 ViewModel
/// </summary>
public class MainViewModel : INotifyPropertyChanged, IDisposable
{
    private RoomData? _selectedRoom;
    private string _searchText = string.Empty;
    private bool _isLoading;

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<RoomData> Rooms { get; } = new();
    public ObservableCollection<RoomData> FilteredRooms { get; } = new();
    public ObservableCollection<LevelGroup> LevelGroups { get; } = new();
    
    // 字段管理器（延迟初始化，避免构造函数异常）
    private FieldManager? _fieldManager;
    public FieldManager FieldManager => _fieldManager ??= new FieldManager();

    // 撤销/重做管理器（延迟初始化）
    private UndoRedoManager? _undoRedoManager;
    public UndoRedoManager UndoRedoManager => _undoRedoManager ??= new UndoRedoManager();

    // 自动保存服务（延迟初始化）
    private AutoSaveService? _autoSaveService;
    private string _savePath = "";
    
    // 事件处理程序引用（用于取消订阅）
    private EventHandler<OperationEventArgs>? _undoRedoHandler;
    private EventHandler<AutoSaveEventArgs>? _autoSavedHandler;
    private EventHandler? _unsavedChangesHandler;

    // 是否有未保存的更改
    private bool _hasUnsavedChanges;
    public bool HasUnsavedChanges
    {
        get => _hasUnsavedChanges;
        set
        {
            _hasUnsavedChanges = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(StatusMessage));
        }
    }

    public RoomData? SelectedRoom
    {
        get => _selectedRoom;
        set
        {
            _selectedRoom = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasSelection));
            RefreshParameterGroups();
        }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            _searchText = value;
            OnPropertyChanged();
            FilterRooms();
        }
    }

    /// <summary>
    /// 楼层名称列表（用于过滤下拉框）
    /// </summary>
    public ObservableCollection<string> LevelNames { get; } = new() { "全部楼层" };
    
    private string _selectedLevel = "全部楼层";
    public string SelectedLevel
    {
        get => _selectedLevel;
        set
        {
            _selectedLevel = value;
            OnPropertyChanged();
            FilterRooms();
        }
    }

    public bool IsLoading
    {
        get => _isLoading;
        set
        {
            _isLoading = value;
            OnPropertyChanged();
        }
    }

    public bool HasSelection => SelectedRoom != null;

    /// <summary>
    /// 当前选中房间的参数分组（供详情面板绑定）
    /// </summary>
    public ObservableCollection<ParameterGroupVM> ParameterGroups { get; } = new();

    private void RefreshParameterGroups()
    {
        ParameterGroups.Clear();
        if (SelectedRoom == null) return;

        var visibleParams = SelectedRoom.AllParameters
            .Where(p => FieldManager.IsFieldVisible(p.Name));

        var groups = visibleParams
            .GroupBy(p => p.GroupName)
            .OrderBy(g => g.Key);

        foreach (var group in groups)
        {
            ParameterGroups.Add(new ParameterGroupVM
            {
                GroupName = group.Key,
                Parameters = new ObservableCollection<RoomParameterInfo>(group.OrderBy(p => p.DisplayName))
            });
        }
    }

    // Excel 服务（延迟初始化，避免构造函数异常）
    private ExcelService? _excelService;

    // 导入导出命令
    public ICommand ExportToExcelCommand { get; }
    public ICommand ImportFromExcelCommand { get; }
    public ICommand ExportTemplateCommand { get; }

    // 撤销/重做命令（延迟初始化）
    private RelayCommand? _undoCommand;
    private RelayCommand? _redoCommand;
    private RelayCommand? _saveCommand;

    public ICommand UndoCommand => _undoCommand ??= new RelayCommand(Undo, () => UndoRedoManager.CanUndo);
    public ICommand RedoCommand => _redoCommand ??= new RelayCommand(Redo, () => UndoRedoManager.CanRedo);
    public ICommand SaveCommand => _saveCommand ??= new RelayCommand(Save);

    // 导入导出状态
    private string _statusMessage = string.Empty;
    public string StatusMessage
    {
        get => _statusMessage;
        set
        {
            _statusMessage = value;
            OnPropertyChanged();
        }
    }

    private int _importProgress;
    public int ImportProgress
    {
        get => _importProgress;
        set
        {
            _importProgress = value;
            OnPropertyChanged();
        }
    }

    // 统计信息
    public int TotalRooms => Rooms.Count;
    public int CompletedRooms => Rooms.Count(r => r.IsComplete);
    public int IncompleteRooms => Rooms.Count(r => !r.IsComplete);
    public double TotalArea => Rooms.Sum(r => r.Area);

    private bool _eventsSubscribed = false;

    public MainViewModel()
    {
        // 初始化 Excel 命令
        ExportToExcelCommand = new RelayCommand(ExportToExcel, CanExport);
        ImportFromExcelCommand = new RelayCommand(ImportFromExcel, CanImport);
        ExportTemplateCommand = new RelayCommand(ExportTemplate, CanExport);

        // 延迟订阅撤销/重做状态变化，避免构造函数异常
        _undoRedoHandler = (s, e) =>
        {
            _undoCommand?.RaiseCanExecuteChanged();
            _redoCommand?.RaiseCanExecuteChanged();
            HasUnsavedChanges = true;
        };
        // 不在构造函数中订阅，改为延迟订阅
    }

    /// <summary>
    /// 确保事件已订阅（延迟订阅）
    /// </summary>
    private void EnsureEventsSubscribed()
    {
        if (!_eventsSubscribed)
        {
            UndoRedoManager.OperationExecuted += _undoRedoHandler;
            _eventsSubscribed = true;
        }
    }

    /// <summary>
    /// 获取 Excel 服务（延迟初始化）
    /// </summary>
    private ExcelService GetExcelService()
    {
        if (_excelService == null)
        {
            try
            {
                _excelService = new ExcelService();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExcelService 初始化失败: {ex.Message}");
                throw;
            }
        }
        return _excelService;
    }

    /// <summary>
    /// 初始化自动保存
    /// </summary>
    public void InitializeAutoSave(string projectPath)
    {
        _savePath = Path.Combine(projectPath, "room_data.json");
        _autoSaveService = new AutoSaveService(Rooms, _savePath, 60);

        _autoSavedHandler = (s, e) =>
        {
            if (e.Success)
            {
                StatusMessage = $"已自动保存 ({e.SaveTime:HH:mm:ss})";
                HasUnsavedChanges = false;
            }
        };
        _autoSaveService.AutoSaved += _autoSavedHandler;

        _unsavedChangesHandler = (s, e) =>
        {
            HasUnsavedChanges = _autoSaveService.HasUnsavedChanges;
        };
        _autoSaveService.UnsavedChangesChanged += _unsavedChangesHandler;
    }

    /// <summary>
    /// 撤销
    /// </summary>
    public void Undo()
    {
        EnsureEventsSubscribed(); // 确保事件已订阅
        var operation = UndoRedoManager.Undo();
        if (operation != null)
        {
            StatusMessage = $"已撤销: {operation.Description}";
        }
    }

    /// <summary>
    /// 重做
    /// </summary>
    public void Redo()
    {
        EnsureEventsSubscribed(); // 确保事件已订阅
        var operation = UndoRedoManager.Redo();
        if (operation != null)
        {
            StatusMessage = $"已重做: {operation.Description}";
        }
    }

    /// <summary>
    /// 保存
    /// </summary>
    public void Save()
    {
        if (_autoSaveService != null)
        {
            if (_autoSaveService.Save())
            {
                StatusMessage = "保存成功";
                HasUnsavedChanges = false;
            }
            else
            {
                StatusMessage = "保存失败";
            }
        }
    }

    /// <summary>
    /// 修改房间名称（支持撤销）
    /// </summary>
    public void RenameRoom(RoomData room, string newName)
    {
        EnsureEventsSubscribed(); // 确保事件已订阅
        var operation = new RenameRoomOperation(room, newName);
        UndoRedoManager.ExecuteOperation(operation);
    }

    /// <summary>
    /// 修改房间编号（支持撤销）
    /// </summary>
    public void RenumberRoom(RoomData room, string newNumber)
    {
        EnsureEventsSubscribed(); // 确保事件已订阅
        var operation = new RenumberRoomOperation(room, newNumber);
        UndoRedoManager.ExecuteOperation(operation);
    }

    /// <summary>
    /// 设置参数（支持撤销）
    /// </summary>
    public void SetParameter(RoomData room, string parameterName, object? value)
    {
        EnsureEventsSubscribed(); // 确保事件已订阅
        var operation = new SetParameterOperation(room, parameterName, value);
        UndoRedoManager.ExecuteOperation(operation);
    }

    public void RefreshStatistics()
    {
        OnPropertyChanged(nameof(TotalRooms));
        OnPropertyChanged(nameof(CompletedRooms));
        OnPropertyChanged(nameof(IncompleteRooms));
        OnPropertyChanged(nameof(TotalArea));
    }

    #region Excel 导入导出

    private bool CanExport() => Rooms.Count > 0 && !IsLoading;
    private bool CanImport() => !IsLoading;

    private void ExportToExcel()
    {
        try
        {
            IsLoading = true;
            StatusMessage = "正在导出...";

            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel 文件|*.xlsx",
                Title = "导出房间数据",
                FileName = $"房间列表_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                GetExcelService().ExportToExcel(Rooms, saveDialog.FileName, includeCustomParams: true);
                StatusMessage = $"导出成功: {Path.GetFileName(saveDialog.FileName)}";
            }
            else
            {
                StatusMessage = "导出已取消";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"导出失败: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    private void ImportFromExcel()
    {
        try
        {
            IsLoading = true;
            StatusMessage = "正在导入...";

            var openDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel 文件|*.xlsx;*.xls",
                Title = "导入房间数据"
            };

            if (openDialog.ShowDialog() == true)
            {
                var updates = GetExcelService().ImportFromExcel(openDialog.FileName);
                
                if (updates.Count == 0)
                {
                    StatusMessage = "未找到有效数据";
                    return;
                }

                // 应用更新
                int successCount = 0;
                int failCount = 0;

                foreach (var update in updates)
                {
                    var room = Rooms.FirstOrDefault(r => r.ElementId == update.ElementId);
                    if (room != null)
                    {
                        if (!string.IsNullOrWhiteSpace(update.Name))
                            room.Name = update.Name;
                        if (!string.IsNullOrWhiteSpace(update.Number))
                            room.Number = update.Number;

                        foreach (var (key, value) in update.CustomParameters)
                        {
                            room.CustomParameters[key] = value;
                        }

                        successCount++;
                    }
                    else
                    {
                        failCount++;
                    }

                    ImportProgress = (int)((double)updates.IndexOf(update) / updates.Count * 100);
                }

                RefreshStatistics();
                StatusMessage = $"导入完成: 成功 {successCount} 条, 失败 {failCount} 条";
            }
            else
            {
                StatusMessage = "导入已取消";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"导入失败: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
            ImportProgress = 0;
        }
    }

    private void ExportTemplate()
    {
        try
        {
            IsLoading = true;
            StatusMessage = "正在生成模板...";

            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel 文件|*.xlsx",
                Title = "导出导入模板",
                FileName = "房间导入模板.xlsx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                var customParams = Rooms
                    .SelectMany(r => r.CustomParameters.Keys)
                    .Distinct()
                    .OrderBy(k => k);

                GetExcelService().ExportTemplate(saveDialog.FileName, Rooms, customParams);
                StatusMessage = $"模板已导出: {Path.GetFileName(saveDialog.FileName)}";
            }
            else
            {
                StatusMessage = "导出已取消";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"导出失败: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    #endregion

    public void LoadRooms(IEnumerable<RoomData> rooms)
    {
        Rooms.Clear();
        foreach (var room in rooms)
        {
            Rooms.Add(room);
        }
        
        // 填充楼层过滤列表
        LevelNames.Clear();
        LevelNames.Add("全部楼层");
        foreach (var level in Rooms.Select(r => r.Level).Distinct().OrderBy(l => l))
        {
            LevelNames.Add(level);
        }
        SelectedLevel = "全部楼层";
        
        GroupByLevel();
        FilterRooms();
        RefreshStatistics();
    }

    private void GroupByLevel()
    {
        LevelGroups.Clear();
        var groups = Rooms
            .GroupBy(r => r.Level)
            .OrderBy(g => g.Key)
            .Select(g => new LevelGroup
            {
                LevelName = g.Key,
                Rooms = new ObservableCollection<RoomData>(g.OrderBy(r => r.Number))
            });

        foreach (var group in groups)
        {
            LevelGroups.Add(group);
        }
    }

    private void FilterRooms()
    {
        FilteredRooms.Clear();
        
        var query = Rooms.AsEnumerable();
        
        // 楼层过滤
        if (!string.IsNullOrEmpty(SelectedLevel) && SelectedLevel != "全部楼层")
        {
            query = query.Where(r => r.Level == SelectedLevel);
        }
        
        // 文字搜索
        if (!string.IsNullOrWhiteSpace(SearchText))
        {
            var search = SearchText.ToLowerInvariant();
            query = query.Where(r => 
                r.Name.ToLowerInvariant().Contains(search) ||
                r.Number.ToLowerInvariant().Contains(search));
        }

        foreach (var room in query)
        {
            FilteredRooms.Add(room);
        }
    }

    public void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public void Dispose()
    {
        try
        {
            // 取消订阅事件
            if (_undoRedoHandler != null && _eventsSubscribed)
            {
                UndoRedoManager.OperationExecuted -= _undoRedoHandler;
                _undoRedoHandler = null;
                _eventsSubscribed = false;
            }
            
            // 停止并释放自动保存服务
            if (_autoSaveService != null)
            {
                if (_autoSavedHandler != null)
                {
                    _autoSaveService.AutoSaved -= _autoSavedHandler;
                    _autoSavedHandler = null;
                }
                if (_unsavedChangesHandler != null)
                {
                    _autoSaveService.UnsavedChangesChanged -= _unsavedChangesHandler;
                    _unsavedChangesHandler = null;
                }
                _autoSaveService.Dispose();
                _autoSaveService = null;
            }
            
            // 清理集合
            Rooms.Clear();
            FilteredRooms.Clear();
            LevelGroups.Clear();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"MainViewModel.Dispose 错误: {ex.Message}");
        }
    }
}

/// <summary>
/// 楼层分组
/// </summary>
public class LevelGroup
{
    public string LevelName { get; set; } = string.Empty;
    public ObservableCollection<RoomData> Rooms { get; set; } = new();
    public int RoomCount => Rooms.Count;
    public bool IsExpanded { get; set; } = true;
}

/// <summary>
/// 参数分组 ViewModel（详情面板用）
/// </summary>
public class ParameterGroupVM
{
    public string GroupName { get; set; } = string.Empty;
    public ObservableCollection<RoomParameterInfo> Parameters { get; set; } = new();
    
    /// <summary>
    /// 是否展开（常用组默认展开）
    /// </summary>
    public bool IsExpanded { get; set; } = true;
}
