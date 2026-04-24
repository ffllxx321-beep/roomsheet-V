namespace RoomManager.Models;

using System.Collections.Generic;
using Newtonsoft.Json;
using System.Windows.Media.Imaging;

/// <summary>
/// 房间数据模型
/// </summary>
public class RoomData
{
    public long ElementId { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Number { get; set; } = string.Empty;
    public string Level { get; set; } = string.Empty;
    public double Area { get; set; }
    public double Volume { get; set; }
    public string Phase { get; set; } = string.Empty;
    
    /// <summary>
    /// 房间类型分类（自动识别）
    /// </summary>
    public RoomCategory Category { get; set; } = RoomCategory.Other;
    
    /// <summary>
    /// 自定义参数（旧版兼容）
    /// </summary>
    public Dictionary<string, object?> CustomParameters { get; set; } = new();

    /// <summary>
    /// 所有 Revit 参数（含只读/可编辑，按组分类）
    /// </summary>
    [JsonIgnore]
    public List<RoomParameterInfo> AllParameters { get; set; } = new();
    
    /// <summary>
    /// 是否已录入完整信息
    /// </summary>
    public bool IsComplete { get; set; }
    
    /// <summary>
    /// 是否选中（用于多选）
    /// </summary>
    [JsonIgnore]
    public bool IsSelected { get; set; }
    
    /// <summary>
    /// 房间略缩图（不序列化）
    /// </summary>
    [JsonIgnore]
    public BitmapImage? Thumbnail { get; set; }
}

/// <summary>
/// 房间类型分类
/// </summary>
public enum RoomCategory
{
    Office,         // 办公室
    MeetingRoom,    // 会议室
    Restroom,       // 卫生间
    Corridor,       // 走廊
    Staircase,      // 楼梯间
    Storage,        // 仓库
    Workshop,       // 车间
    Kitchen,        // 厨房
    Bedroom,        // 卧室
    LivingRoom,     // 客厅
    Other           // 其他
}

/// <summary>
/// 房间类型映射规则
/// </summary>
public static class RoomCategoryHelper
{
    private static readonly Dictionary<RoomCategory, string[]> Keywords = new()
    {
        [RoomCategory.Office] = new[] { "办公室", "办公", "OFFICE", "OFF" },
        [RoomCategory.MeetingRoom] = new[] { "会议室", "会议", "MEETING", "MEET" },
        [RoomCategory.Restroom] = new[] { "卫生间", "厕所", "洗手间", "RESTROOM", "WC", "TOILET" },
        [RoomCategory.Corridor] = new[] { "走廊", "过道", "CORRIDOR", "HALLWAY" },
        [RoomCategory.Staircase] = new[] { "楼梯", "楼梯间", "STAIR", "STAIRCASE" },
        [RoomCategory.Storage] = new[] { "仓库", "储藏", "STORAGE", "STORE" },
        [RoomCategory.Workshop] = new[] { "车间", "工厂", "WORKSHOP", "FACTORY" },
        [RoomCategory.Kitchen] = new[] { "厨房", "KITCHEN" },
        [RoomCategory.Bedroom] = new[] { "卧室", "BEDROOM", "卧室" },
        [RoomCategory.LivingRoom] = new[] { "客厅", "LIVING", "起居" },
    };

    public static RoomCategory Classify(string roomName)
    {
        if (string.IsNullOrWhiteSpace(roomName))
            return RoomCategory.Other;

        var upperName = roomName.ToUpperInvariant();
        
        foreach (var (category, keywords) in Keywords)
        {
            if (keywords.Any(kw => upperName.Contains(kw.ToUpperInvariant())))
                return category;
        }

        return RoomCategory.Other;
    }

    public static string GetDisplayName(RoomCategory category)
    {
        return category switch
        {
            RoomCategory.Office => "办公室",
            RoomCategory.MeetingRoom => "会议室",
            RoomCategory.Restroom => "卫生间",
            RoomCategory.Corridor => "走廊",
            RoomCategory.Staircase => "楼梯间",
            RoomCategory.Storage => "仓库",
            RoomCategory.Workshop => "车间",
            RoomCategory.Kitchen => "厨房",
            RoomCategory.Bedroom => "卧室",
            RoomCategory.LivingRoom => "客厅",
            _ => "其他"
        };
    }
}

/// <summary>
/// Revit 房间参数信息
/// </summary>
public class RoomParameterInfo : System.ComponentModel.INotifyPropertyChanged
{
    public string Name { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string GroupName { get; set; } = string.Empty;
    public bool IsReadOnly { get; set; }
    public string StorageTypeName { get; set; } = string.Empty;
    
    /// <summary>
    /// 原始值（用于检测修改）
    /// </summary>
    public string OriginalValue { get; set; } = string.Empty;
    
    /// <summary>
    /// 是否被修改过
    /// </summary>
    public bool IsModified => !IsReadOnly && DisplayValue != OriginalValue;
    
    private string _displayValue = string.Empty;
    public string DisplayValue
    {
        get => _displayValue;
        set
        {
            if (_displayValue != value)
            {
                _displayValue = value;
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(DisplayValue)));
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsModified)));
            }
        }
    }

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
}
