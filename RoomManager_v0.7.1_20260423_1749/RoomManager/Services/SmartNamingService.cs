using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace RoomManager.Services;

/// <summary>
/// 智能命名建议服务
/// </summary>
public class SmartNamingService
{
    private readonly Document _document;

    public SmartNamingService(Document document)
    {
        _document = document;
    }

    /// <summary>
    /// 根据相邻房间推断名称
    /// </summary>
    public List<NamingSuggestion> SuggestNameFromNeighbors(Room room)
    {
        var suggestions = new List<NamingSuggestion>();

        // 获取相邻房间
        var neighbors = GetAdjacentRooms(room);
        if (neighbors.Count == 0) return suggestions;

        // 分析相邻房间名称模式
        var patterns = AnalyzeNamingPatterns(neighbors);

        foreach (var pattern in patterns.Take(3))
        {
            suggestions.Add(new NamingSuggestion
            {
                SuggestedName = GenerateNameFromPattern(pattern, room),
                Confidence = pattern.Confidence,
                Reason = $"基于 {pattern.MatchCount} 个相邻房间的命名模式",
                Pattern = pattern
            });
        }

        return suggestions;
    }

    /// <summary>
    /// 根据面积推断房间类型
    /// </summary>
    public List<NamingSuggestion> SuggestNameFromArea(Room room)
    {
        var suggestions = new List<NamingSuggestion>();
        var area = room.Area;

        // 面积规则（单位：平方英尺，Revit 默认单位）
        // 1 m² ≈ 10.764 sq ft
        var areaRules = new List<(double minArea, double maxArea, string name, double confidence)>
        {
            (0, 50, "储藏室", 0.6),
            (50, 150, "卫生间", 0.7),
            (150, 300, "办公室", 0.8),
            (300, 600, "会议室", 0.75),
            (600, 1500, "大会议室", 0.7),
            (1500, 3000, "多功能厅", 0.65),
            (3000, double.MaxValue, "大厅", 0.6)
        };

        foreach (var (minArea, maxArea, name, confidence) in areaRules)
        {
            if (area >= minArea && area < maxArea)
            {
                suggestions.Add(new NamingSuggestion
                {
                    SuggestedName = name,
                    Confidence = confidence,
                    Reason = $"基于面积推断 ({area / 10.764:F1} m²)",
                    Pattern = new NamingPattern { Type = PatternType.Area }
                });
            }
        }

        return suggestions;
    }

    /// <summary>
    /// 根据楼层和编号模式建议名称
    /// </summary>
    public List<NamingSuggestion> SuggestNameFromFloorAndNumber(Room room)
    {
        var suggestions = new List<NamingSuggestion>();
        var level = room.Level;
        if (level == null) return suggestions;

        var floorName = level.Name;
        var floorNumber = ExtractFloorNumber(floorName);
        var roomNumber = room.Number;

        // 提取编号中的序号
        var indexMatch = Regex.Match(roomNumber, @"\d+");
        if (!indexMatch.Success) return suggestions;

        var index = int.Parse(indexMatch.Value);

        // 常见命名模板
        var templates = new List<(string template, string description, double confidence)>
        {
            ("办公室-{floor}-{index:00}", "标准办公室命名", 0.7),
            ("会议室-{floor}-{index:00}", "标准会议室命名", 0.65),
            ("房间-{floor}-{index:00}", "通用房间命名", 0.5)
        };

        foreach (var (template, description, confidence) in templates)
        {
            var suggestedName = template
                .Replace("{floor}", floorNumber)
                .Replace("{index:00}", index.ToString("00"))
                .Replace("{index}", index.ToString());

            suggestions.Add(new NamingSuggestion
            {
                SuggestedName = suggestedName,
                Confidence = confidence,
                Reason = description,
                Pattern = new NamingPattern { Type = PatternType.Template, Template = template }
            });
        }

        return suggestions;
    }

    /// <summary>
    /// 综合建议（合并所有来源）
    /// </summary>
    public List<NamingSuggestion> GetComprehensiveSuggestions(Room room)
    {
        var allSuggestions = new List<NamingSuggestion>();

        // 从相邻房间推断
        allSuggestions.AddRange(SuggestNameFromNeighbors(room));

        // 从面积推断
        allSuggestions.AddRange(SuggestNameFromArea(room));

        // 从楼层编号推断
        allSuggestions.AddRange(SuggestNameFromFloorAndNumber(room));

        // 去重并排序
        return allSuggestions
            .GroupBy(s => s.SuggestedName)
            .Select(g => g.OrderByDescending(s => s.Confidence).First())
            .OrderByDescending(s => s.Confidence)
            .Take(5)
            .ToList();
    }

    /// <summary>
    /// 批量生成命名建议
    /// </summary>
    public Dictionary<long, List<NamingSuggestion>> BatchSuggestNames(IEnumerable<Room> rooms)
    {
        var results = new Dictionary<long, List<NamingSuggestion>>();

        foreach (var room in rooms)
        {
            results[room.Id.Value] = GetComprehensiveSuggestions(room);
        }

        return results;
    }

    /// <summary>
    /// 应用命名建议
    /// </summary>
    public bool ApplySuggestion(Room room, NamingSuggestion suggestion)
    {
        try
        {
            using var transaction = new Transaction(_document, "应用命名建议");
            transaction.Start();

            room.Name = suggestion.SuggestedName;

            transaction.Commit();
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 批量应用命名建议
    /// </summary>
    public Dictionary<long, bool> BatchApplySuggestions(Dictionary<long, NamingSuggestion> suggestions)
    {
        var results = new Dictionary<long, bool>();

        using var transaction = new Transaction(_document, "批量应用命名建议");
        transaction.Start();

        try
        {
            foreach (var (roomId, suggestion) in suggestions)
            {
                try
                {
                    var room = _document.GetElement(new ElementId(roomId)) as Room;
                    if (room != null)
                    {
                        room.Name = suggestion.SuggestedName;
                        results[roomId] = true;
                    }
                }
                catch
                {
                    results[roomId] = false;
                }
            }

            transaction.Commit();
        }
        catch
        {
            transaction.RollBack();
        }

        return results;
    }

    /// <summary>
    /// 获取相邻房间
    /// </summary>
    private List<Room> GetAdjacentRooms(Room room)
    {
        var adjacentRooms = new List<Room>();

        try
        {
            // 获取房间边界
            var boundarySegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
            if (boundarySegments == null || boundarySegments.Count == 0)
                return adjacentRooms;

            // 收集所有边界元素
            var boundaryElementIds = new HashSet<long>();
            foreach (var loop in boundarySegments)
            {
                foreach (var segment in loop)
                {
                    var elementId = segment.ElementId;
                    if (elementId != ElementId.InvalidElementId)
                    {
                        boundaryElementIds.Add(elementId.Value);
                    }
                }
            }

            // 查找共享边界的房间
            var allRooms = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r.Id.Value != room.Id.Value);

            foreach (var otherRoom in allRooms)
            {
                var otherBoundary = otherRoom.GetBoundarySegments(new SpatialElementBoundaryOptions());
                if (otherBoundary == null) continue;

                foreach (var loop in otherBoundary)
                {
                    foreach (var segment in loop)
                    {
                        if (boundaryElementIds.Contains(segment.ElementId.Value))
                        {
                            adjacentRooms.Add(otherRoom);
                            break;
                        }
                    }
                    if (adjacentRooms.Contains(otherRoom)) break;
                }
            }
        }
        catch
        {
            // 忽略错误
        }

        return adjacentRooms;
    }

    /// <summary>
    /// 分析命名模式
    /// </summary>
    private List<NamingPattern> AnalyzeNamingPatterns(List<Room> rooms)
    {
        var patterns = new List<NamingPattern>();

        // 提取所有房间名称
        var names = rooms
            .Where(r => !string.IsNullOrWhiteSpace(r.Name))
            .Select(r => r.Name)
            .ToList();

        if (names.Count == 0) return patterns;

        // 分析前缀模式
        var prefixGroups = names
            .GroupBy(n => GetPrefix(n))
            .Where(g => g.Count() >= 2)
            .OrderByDescending(g => g.Count());

        foreach (var group in prefixGroups)
        {
            if (string.IsNullOrEmpty(group.Key)) continue;

            patterns.Add(new NamingPattern
            {
                Type = PatternType.Prefix,
                Prefix = group.Key,
                MatchCount = group.Count(),
                Confidence = Math.Min(0.9, group.Count() * 0.2)
            });
        }

        // 分析后缀模式
        var suffixGroups = names
            .GroupBy(n => GetSuffix(n))
            .Where(g => g.Count() >= 2)
            .OrderByDescending(g => g.Count());

        foreach (var group in suffixGroups)
        {
            if (string.IsNullOrEmpty(group.Key)) continue;

            patterns.Add(new NamingPattern
            {
                Type = PatternType.Suffix,
                Suffix = group.Key,
                MatchCount = group.Count(),
                Confidence = Math.Min(0.85, group.Count() * 0.18)
            });
        }

        // 分析类型模式
        var typeGroups = names
            .GroupBy(n => ExtractRoomType(n))
            .Where(g => g.Count() >= 2)
            .OrderByDescending(g => g.Count());

        foreach (var group in typeGroups)
        {
            if (string.IsNullOrEmpty(group.Key)) continue;

            patterns.Add(new NamingPattern
            {
                Type = PatternType.RoomType,
                RoomType = group.Key,
                MatchCount = group.Count(),
                Confidence = Math.Min(0.8, group.Count() * 0.15)
            });
        }

        return patterns;
    }

    /// <summary>
    /// 根据模式生成名称
    /// </summary>
    private string GenerateNameFromPattern(NamingPattern pattern, Room room)
    {
        return pattern.Type switch
        {
            PatternType.Prefix => $"{pattern.Prefix}{room.Number}",
            PatternType.Suffix => $"{room.Number}{pattern.Suffix}",
            PatternType.RoomType => $"{pattern.RoomType}-{room.Number}",
            PatternType.Template => pattern.Template
                .Replace("{floor}", ExtractFloorNumber(room.Level?.Name ?? ""))
                .Replace("{index:00}", ExtractIndex(room.Number).ToString("00")),
            _ => room.Name
        };
    }

    /// <summary>
    /// 获取前缀
    /// </summary>
    private string GetPrefix(string name)
    {
        var match = Regex.Match(name, @"^[^\d\-]+");
        return match.Success ? match.Value.Trim() : "";
    }

    /// <summary>
    /// 获取后缀
    /// </summary>
    private string GetSuffix(string name)
    {
        var match = Regex.Match(name, @"[^\d\-]+$");
        return match.Success ? match.Value.Trim() : "";
    }

    /// <summary>
    /// 提取房间类型
    /// </summary>
    private string ExtractRoomType(string name)
    {
        var types = new[] { "办公室", "会议室", "卫生间", "走廊", "楼梯", "储藏", "厨房", "卧室", "客厅" };
        foreach (var type in types)
        {
            if (name.Contains(type)) return type;
        }
        return "";
    }

    /// <summary>
    /// 提取楼层编号
    /// </summary>
    private string ExtractFloorNumber(string levelName)
    {
        var match = Regex.Match(levelName ?? "", @"\d+");
        return match.Success ? match.Value : "0";
    }

    /// <summary>
    /// 提取序号
    /// </summary>
    private int ExtractIndex(string roomNumber)
    {
        var match = Regex.Match(roomNumber ?? "", @"\d+");
        return match.Success ? int.Parse(match.Value) : 1;
    }
}

/// <summary>
/// 命名建议
/// </summary>
public class NamingSuggestion
{
    public string SuggestedName { get; set; } = string.Empty;
    public double Confidence { get; set; }
    public string Reason { get; set; } = string.Empty;
    public NamingPattern? Pattern { get; set; }
}

/// <summary>
/// 命名模式
/// </summary>
public class NamingPattern
{
    public PatternType Type { get; set; }
    public string Prefix { get; set; } = string.Empty;
    public string Suffix { get; set; } = string.Empty;
    public string RoomType { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
    public int MatchCount { get; set; }
    public double Confidence { get; set; }
}

/// <summary>
/// 模式类型
/// </summary>
public enum PatternType
{
    Prefix,
    Suffix,
    RoomType,
    Template,
    Area
}
