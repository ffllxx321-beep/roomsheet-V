using RoomManager.Models;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace RoomManager.Services;

/// <summary>
/// 自动保存服务
/// </summary>
public class AutoSaveService : IDisposable
{
    private readonly string _savePath;
    private readonly int _autoSaveInterval;
    private readonly ObservableCollection<RoomData> _rooms;
    private System.Timers.Timer? _autoSaveTimer;
    private bool _hasUnsavedChanges = false;
    private bool _disposed = false;

    /// <summary>
    /// 是否有未保存的更改
    /// </summary>
    public bool HasUnsavedChanges
    {
        get => _hasUnsavedChanges;
        set
        {
            if (_hasUnsavedChanges != value)
            {
                _hasUnsavedChanges = value;
                UnsavedChangesChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    /// <summary>
    /// 上次保存时间
    /// </summary>
    public DateTime LastSaveTime { get; private set; }

    /// <summary>
    /// 自动保存事件
    /// </summary>
    public event EventHandler<AutoSaveEventArgs>? AutoSaved;

    /// <summary>
    /// 未保存状态变化事件
    /// </summary>
    public event EventHandler? UnsavedChangesChanged;

    public AutoSaveService(ObservableCollection<RoomData> rooms, string savePath, int autoSaveIntervalSeconds = 60)
    {
        _rooms = rooms;
        _savePath = savePath;
        _autoSaveInterval = autoSaveIntervalSeconds;

        // 确保目录存在
        var directory = Path.GetDirectoryName(_savePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        StartAutoSave();
    }

    /// <summary>
    /// 启动自动保存
    /// </summary>
    public void StartAutoSave()
    {
        if (_autoSaveTimer != null) return;

        _autoSaveTimer = new System.Timers.Timer(_autoSaveInterval * 1000);
        _autoSaveTimer.Elapsed += OnAutoSaveTimer;
        _autoSaveTimer.Start();
    }

    /// <summary>
    /// 停止自动保存
    /// </summary>
    public void StopAutoSave()
    {
        if (_autoSaveTimer == null) return;
        
        try
        {
            _autoSaveTimer.Stop();
            _autoSaveTimer.Elapsed -= OnAutoSaveTimer;
            _autoSaveTimer.Dispose();
            _autoSaveTimer = null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"StopAutoSave 错误: {ex.Message}");
        }
    }

    private void OnAutoSaveTimer(object? sender, System.Timers.ElapsedEventArgs e)
    {
        if (_disposed) return;
        
        try
        {
            if (HasUnsavedChanges)
            {
                Save();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"OnAutoSaveTimer 错误: {ex.Message}");
        }
    }

    /// <summary>
    /// 标记有更改
    /// </summary>
    public void MarkDirty()
    {
        HasUnsavedChanges = true;
    }

    /// <summary>
    /// 保存数据
    /// </summary>
    public bool Save()
    {
        try
        {
            var json = JsonConvert.SerializeObject(_rooms, Formatting.Indented);

            // 先写入临时文件，再替换（防止写入失败导致数据丢失）
            var tempPath = _savePath + ".tmp";
            File.WriteAllText(tempPath, json);
            File.Move(tempPath, _savePath, true);

            LastSaveTime = DateTime.Now;
            HasUnsavedChanges = false;

            AutoSaved?.Invoke(this, new AutoSaveEventArgs
            {
                Success = true,
                SaveTime = LastSaveTime,
                RoomCount = _rooms.Count
            });

            return true;
        }
        catch (Exception ex)
        {
            AutoSaved?.Invoke(this, new AutoSaveEventArgs
            {
                Success = false,
                Error = ex.Message
            });

            return false;
        }
    }

    /// <summary>
    /// 加载数据
    /// </summary>
    public List<RoomData>? Load()
    {
        try
        {
            if (!File.Exists(_savePath)) return null;

            var json = File.ReadAllText(_savePath);
            var rooms = JsonConvert.DeserializeObject<List<RoomData>>(json);

            HasUnsavedChanges = false;

            return rooms;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 创建备份
    /// </summary>
    public string CreateBackup()
    {
        var backupPath = _savePath.Replace(".json", $"_backup_{DateTime.Now:yyyyMMdd_HHmmss}.json");
        
        try
        {
            if (File.Exists(_savePath))
            {
                File.Copy(_savePath, backupPath);
            }
            return backupPath;
        }
        catch
        {
            return "";
        }
    }

    /// <summary>
    /// 获取所有备份文件
    /// </summary>
    public List<BackupInfo> GetBackups()
    {
        var directory = Path.GetDirectoryName(_savePath) ?? "";
        var fileName = Path.GetFileNameWithoutExtension(_savePath);

        if (!Directory.Exists(directory)) return new();

        var backups = Directory.GetFiles(directory, $"{fileName}_backup_*.json")
            .Select(f => new BackupInfo
            {
                FilePath = f,
                FileName = Path.GetFileName(f),
                CreateTime = File.GetCreationTime(f),
                Size = new FileInfo(f).Length
            })
            .OrderByDescending(b => b.CreateTime)
            .ToList();

        return backups;
    }

    /// <summary>
    /// 从备份恢复
    /// </summary>
    public bool RestoreFromBackup(string backupPath)
    {
        try
        {
            if (!File.Exists(backupPath)) return false;

            var json = File.ReadAllText(backupPath);
            var rooms = JsonConvert.DeserializeObject<List<RoomData>>(json);

            if (rooms == null) return false;

            _rooms.Clear();
            foreach (var room in rooms)
            {
                _rooms.Add(room);
            }

            HasUnsavedChanges = true;

            return true;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            StopAutoSave();
            _disposed = true;
        }
    }
}

/// <summary>
/// 自动保存事件参数
/// </summary>
public class AutoSaveEventArgs : EventArgs
{
    public bool Success { get; set; }
    public DateTime SaveTime { get; set; }
    public int RoomCount { get; set; }
    public string? Error { get; set; }
}

/// <summary>
/// 备份信息
/// </summary>
public class BackupInfo
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public DateTime CreateTime { get; set; }
    public long Size { get; set; }
}
