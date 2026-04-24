namespace RoomManager.Models;

using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

/// <summary>
/// 字段可见性配置项（用于字段管理器显示）
/// </summary>
public class FieldVisibilityItem : INotifyPropertyChanged
{
    public string Name { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string GroupName { get; set; } = string.Empty;
    public bool IsReadOnly { get; set; }

    private bool _isVisible = true;
    public bool IsVisible
    {
        get => _isVisible;
        set
        {
            if (_isVisible != value)
            {
                _isVisible = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsVisible)));
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

/// <summary>
/// 字段管理器 — 管理参数的显示/隐藏
/// </summary>
public class FieldManager
{
    /// <summary>
    /// 隐藏的字段名集合
    /// </summary>
    public HashSet<string> HiddenFields { get; private set; } = new();

    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RoomManager");

    private static readonly string ConfigPath = Path.Combine(ConfigDir, "field_visibility.json");

    public FieldManager()
    {
        LoadConfig();
    }

    /// <summary>
    /// 判断字段是否可见
    /// </summary>
    public bool IsFieldVisible(string fieldName)
    {
        return !HiddenFields.Contains(fieldName);
    }

    /// <summary>
    /// 设置字段可见性
    /// </summary>
    public void SetFieldVisibility(string fieldName, bool visible)
    {
        if (visible)
            HiddenFields.Remove(fieldName);
        else
            HiddenFields.Add(fieldName);
    }

    /// <summary>
    /// 重置所有字段为可见
    /// </summary>
    public void ResetAll()
    {
        HiddenFields.Clear();
    }

    /// <summary>
    /// 保存配置到文件
    /// </summary>
    public void SaveConfig()
    {
        try
        {
            if (!Directory.Exists(ConfigDir))
                Directory.CreateDirectory(ConfigDir);

            var json = JsonConvert.SerializeObject(HiddenFields.ToList(), Formatting.Indented);
            File.WriteAllText(ConfigPath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"保存字段配置失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 从文件加载配置
    /// </summary>
    private void LoadConfig()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                var list = JsonConvert.DeserializeObject<List<string>>(json);
                if (list != null)
                    HiddenFields = new HashSet<string>(list);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"加载字段配置失败: {ex.Message}");
        }
    }

    // === 以下为旧接口兼容，不再使用 ===

    public IReadOnlyList<FieldVisibilityItem> Fields => new List<FieldVisibilityItem>().AsReadOnly();

    public void AddCustomField(string name, FieldType type, List<string>? options = null) { }
    public void RemoveField(string name) { }
    public void ReorderFields(List<string> fieldNames) { }
}

/// <summary>
/// 字段类型（旧版兼容）
/// </summary>
public enum FieldType
{
    Text, Number, Area, Volume, Dropdown, Date, Boolean, MultiLineText
}
