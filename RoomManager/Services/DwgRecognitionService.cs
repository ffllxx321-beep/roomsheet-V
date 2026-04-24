using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DwgLine = ACadSharp.Entities.Line;

namespace RoomManager.Services;

/// <summary>
/// DWG 文字识别服务
/// </summary>
public class DwgRecognitionService
{
    /// <summary>
    /// 从 DWG 文件提取所有文字
    /// </summary>
    public List<DwgTextInfo> ExtractTextsFromDwg(string filePath, string[]? layerFilter = null)
    {
        var texts = new List<DwgTextInfo>();

        try
        {
            // 使用 ACadSharp 读取 DWG（官方文档：DwgReader.Read）
            using var reader = new DwgReader(filePath);
            var document = reader.Read();
            
            foreach (var entity in document.Entities)
            {
                // 图层过滤
                var layerName = "";
                if (entity is TextEntity text)
                {
                    layerName = text.Layer?.Name ?? "";
                    if (layerFilter != null && !layerFilter.Contains(layerName))
                        continue;

                    texts.Add(new DwgTextInfo
                    {
                        Content = text.Value,
                        X = text.InsertPoint.X,
                        Y = text.InsertPoint.Y,
                        Z = text.InsertPoint.Z,
                        Height = text.Height,
                        Rotation = text.Rotation,
                        LayerName = layerName
                    });
                }
                else if (entity is MText mText)
                {
                    layerName = mText.Layer?.Name ?? "";
                    if (layerFilter != null && !layerFilter.Contains(layerName))
                        continue;

                    texts.Add(new DwgTextInfo
                    {
                        Content = mText.Value,
                        X = mText.InsertPoint.X,
                        Y = mText.InsertPoint.Y,
                        Z = mText.InsertPoint.Z,
                        Height = mText.Height,
                        Rotation = mText.Rotation,
                        LayerName = layerName
                    });
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"读取 DWG 文件失败: {ex.Message}");
        }

        return texts;
    }

    /// <summary>
    /// 从 DWG 提取闭合区域（用于房间边界识别）
    /// </summary>
    public List<ClosedRegion> ExtractClosedRegions(string filePath, string[]? layerFilter = null)
    {
        var regions = new List<ClosedRegion>();

        try
        {
            // 使用 ACadSharp 读取 DWG（官方文档：DwgReader.Read）
            using var reader = new DwgReader(filePath);
            var document = reader.Read();
            var curves = new List<DwgCurveInfo>();

            // 提取所有曲线（线、多段线、圆等）
            foreach (var entity in document.Entities)
            {
                var layerName = "";
                
                if (entity is DwgLine line)
                {
                    layerName = line.Layer?.Name ?? "";
                    if (layerFilter != null && !layerFilter.Contains(layerName))
                        continue;

                    curves.Add(new DwgCurveInfo
                    {
                        Type = CurveType.Line,
                        StartX = line.StartPoint.X,
                        StartY = line.StartPoint.Y,
                        EndX = line.EndPoint.X,
                        EndY = line.EndPoint.Y,
                        LayerName = layerName
                    });
                }
                else if (entity is LwPolyline lwPoly)
                {
                    layerName = lwPoly.Layer?.Name ?? "";
                    if (layerFilter != null && !layerFilter.Contains(layerName))
                        continue;

                    // 将多段线分解为线段
                    var vertices = lwPoly.Vertices.ToList();
                    for (int i = 0; i < vertices.Count - 1; i++)
                    {
                        curves.Add(new DwgCurveInfo
                        {
                            Type = CurveType.Line,
                            StartX = vertices[i].Location.X,
                            StartY = vertices[i].Location.Y,
                            EndX = vertices[i + 1].Location.X,
                            EndY = vertices[i + 1].Location.Y,
                            LayerName = layerName
                        });
                    }

                    // 闭合多段线
                    if (lwPoly.IsClosed && vertices.Count > 2)
                    {
                        curves.Add(new DwgCurveInfo
                        {
                            Type = CurveType.Line,
                            StartX = vertices[^1].Location.X,
                            StartY = vertices[^1].Location.Y,
                            EndX = vertices[0].Location.X,
                            EndY = vertices[0].Location.Y,
                            LayerName = layerName
                        });
                    }
                }
                else if (entity is Circle circle)
                {
                    layerName = circle.Layer?.Name ?? "";
                    if (layerFilter != null && !layerFilter.Contains(layerName))
                        continue;

                    curves.Add(new DwgCurveInfo
                    {
                        Type = CurveType.Circle,
                        CenterX = circle.Center.X,
                        CenterY = circle.Center.Y,
                        Radius = circle.Radius,
                        LayerName = layerName
                    });
                }
            }

            // 检测闭合区域
            regions = DetectClosedRegions(curves);
        }
        catch (Exception ex)
        {
            throw new Exception($"提取闭合区域失败: {ex.Message}");
        }

        return regions;
    }

    /// <summary>
    /// 检测闭合区域（基于曲线连接）
    /// </summary>
    private List<ClosedRegion> DetectClosedRegions(List<DwgCurveInfo> curves)
    {
        var regions = new List<ClosedRegion>();
        var usedCurves = new HashSet<int>();

        // 构建曲线连接图
        var graph = BuildCurveGraph(curves);

        // 查找所有闭合环
        for (int i = 0; i < curves.Count; i++)
        {
            if (usedCurves.Contains(i)) continue;

            var loop = FindClosedLoop(curves, graph, i, usedCurves);
            if (loop != null && loop.Count >= 3)
            {
                var region = new ClosedRegion
                {
                    Curves = loop,
                    CenterX = loop.Average(c => (c.StartX + c.EndX) / 2),
                    CenterY = loop.Average(c => (c.StartY + c.EndY) / 2),
                    Area = CalculatePolygonArea(loop)
                };
                regions.Add(region);
            }
        }

        return regions;
    }

    /// <summary>
    /// 构建曲线连接图
    /// </summary>
    private Dictionary<int, List<int>> BuildCurveGraph(List<DwgCurveInfo> curves)
    {
        var graph = new Dictionary<int, List<int>>();
        double tolerance = 0.001; // 连接容差

        for (int i = 0; i < curves.Count; i++)
        {
            graph[i] = new List<int>();

            for (int j = 0; j < curves.Count; j++)
            {
                if (i == j) continue;

                // 检查曲线 i 的终点是否连接到曲线 j 的起点
                var dx = curves[i].EndX - curves[j].StartX;
                var dy = curves[i].EndY - curves[j].StartY;
                var distance = Math.Sqrt(dx * dx + dy * dy);

                if (distance < tolerance)
                {
                    graph[i].Add(j);
                }
            }
        }

        return graph;
    }

    /// <summary>
    /// 查找闭合环
    /// </summary>
    private List<DwgCurveInfo>? FindClosedLoop(
        List<DwgCurveInfo> curves,
        Dictionary<int, List<int>> graph,
        int startIndex,
        HashSet<int> usedCurves)
    {
        var loop = new List<DwgCurveInfo>();
        var visited = new HashSet<int>();
        var current = startIndex;
        int maxIterations = curves.Count;

        while (maxIterations-- > 0)
        {
            if (visited.Contains(current))
            {
                // 找到闭合环
                if (current == startIndex && loop.Count >= 3)
                {
                    foreach (var idx in visited)
                    {
                        usedCurves.Add(idx);
                    }
                    return loop;
                }
                break;
            }

            visited.Add(current);
            loop.Add(curves[current]);

            // 查找下一个连接的曲线
            var nextCurves = graph.GetValueOrDefault(current, new List<int>());
            var next = nextCurves.FirstOrDefault(n => !visited.Contains(n) || n == startIndex);

            if (next == 0 && !nextCurves.Contains(0))
            {
                break;
            }

            current = next;
        }

        return null;
    }

    /// <summary>
    /// 计算多边形面积
    /// </summary>
    private double CalculatePolygonArea(List<DwgCurveInfo> curves)
    {
        if (curves.Count < 3) return 0;

        double area = 0;
        for (int i = 0; i < curves.Count; i++)
        {
            var j = (i + 1) % curves.Count;
            area += curves[i].StartX * curves[j].StartY;
            area -= curves[j].StartX * curves[i].StartY;
        }

        return Math.Abs(area) / 2;
    }

    /// <summary>
    /// 将 DWG 坐标转换为 Revit 坐标
    /// </summary>
    public List<TextInRevit> ConvertToRevitCoordinates(
        List<DwgTextInfo> dwgTexts, 
        Transform transform)
    {
        var revitTexts = new List<TextInRevit>();

        foreach (var dwgText in dwgTexts)
        {
            // 创建 DWG 坐标点
            var dwgPoint = new XYZ(dwgText.X, dwgText.Y, dwgText.Z);
            
            // 应用 Revit 变换矩阵
            var revitPoint = transform.OfPoint(dwgPoint);

            revitTexts.Add(new TextInRevit
            {
                Content = dwgText.Content,
                Position = revitPoint,
                Height = dwgText.Height,
                LayerName = dwgText.LayerName
            });
        }

        return revitTexts;
    }

    /// <summary>
    /// 匹配文字到房间（增强版）
    /// </summary>
    /// <summary>
    /// 匹配文字到房间（最近距离算法 + 双向独占匹配）
    /// </summary>
    public List<RoomTextMatch> MatchTextsToRooms(
        List<TextInRevit> texts, 
        List<Room> rooms,
        double maxDistance = 30.0) // 约 10 米，足够覆盖大多数房间
    {
        var matches = new List<RoomTextMatch>();
        if (texts.Count == 0 || rooms.Count == 0) return matches;

        // 第一步：计算候选对（距离 + 文义 + 是否在房间内）
        var candidates = new List<RoomTextCandidate>();

        foreach (var room in rooms)
        {
            var roomCenter = GetRoomCenter(room);
            if (roomCenter == null) continue;
            var roomNameNormalized = NormalizeForMatch(room.Name);
            var roomNumberNormalized = NormalizeForMatch(room.Number);
            var roomKeywordGroup = PresetKeywordMatcher.GetKeywordGroup(roomNameNormalized);

            foreach (var text in texts)
            {
                if (string.IsNullOrWhiteSpace(text.Content)) continue;
                var normalizedText = NormalizeForMatch(text.Content);
                if (string.IsNullOrWhiteSpace(normalizedText)) continue;

                var dx = text.Position.X - roomCenter.X;
                var dy = text.Position.Y - roomCenter.Y;
                var dist = Math.Sqrt(dx * dx + dy * dy);
                var textPoint = new XYZ(text.Position.X, text.Position.Y, roomCenter.Z);
                var insideRoom = IsPointInRoom(textPoint, room);

                if (!insideRoom && dist > maxDistance) continue;

                var distanceScore = Math.Max(0.0, 1.0 - dist / maxDistance);
                var numberScore = ComputeTextSimilarity(normalizedText, roomNumberNormalized);
                var nameScore = ComputeTextSimilarity(normalizedText, roomNameNormalized);
                var keywordGroup = PresetKeywordMatcher.GetKeywordGroup(normalizedText);
                var keywordScore = !string.IsNullOrWhiteSpace(keywordGroup) && keywordGroup == roomKeywordGroup ? 1.0 : 0.0;

                // 综合得分：优先“文字在房间内”，其次看距离，再看文字/房间名的模糊匹配
                var semanticScore = Math.Max(numberScore, nameScore * 0.8 + keywordScore * 0.2);
                var totalScore = (insideRoom ? 0.45 : 0.0) + distanceScore * 0.35 + semanticScore * 0.20;

                candidates.Add(new RoomTextCandidate
                {
                    Room = room,
                    Text = text,
                    Distance = dist,
                    IsInsideRoom = insideRoom,
                    Score = totalScore
                });
            }
        }

        // 第二步：贪心独占匹配 — 按总得分排序，每个文字/房间只能匹配一次
        candidates.Sort((a, b) => b.Score.CompareTo(a.Score));
        var usedTexts = new HashSet<string>(); // 已被占用的文字（按内容+位置唯一标识）
        var matchedRooms = new Dictionary<long, RoomTextCandidate>();

        foreach (var candidate in candidates)
        {
            var roomId = candidate.Room.Id.Value;
            var textKey = $"{NormalizeForMatch(candidate.Text.Content)}@{candidate.Text.Position.X:F2},{candidate.Text.Position.Y:F2}";
            if (usedTexts.Contains(textKey)) continue;
            if (matchedRooms.ContainsKey(roomId)) continue;
            if (candidate.Score < 0.25) continue; // 过滤低质量匹配

            matchedRooms[roomId] = candidate;
            usedTexts.Add(textKey);
        }

        // 第三步：生成结果
        foreach (var room in rooms)
        {
            var roomId = room.Id.Value;
            TextInRevit? bestMatch = null;
            double confidence = 0.0;
            MatchStatus status = MatchStatus.NoText;
            var alternatives = new List<string>();

            if (matchedRooms.TryGetValue(roomId, out var matched))
            {
                bestMatch = matched.Text;
                confidence = Math.Max(0.1, Math.Min(1.0, matched.Score));
                status = matched.IsInsideRoom || matched.Distance <= 8.0 ? MatchStatus.Matched : MatchStatus.LowConfidence;

                // 备选结果：按综合分数排序
                alternatives = candidates
                    .Where(item => item.Room.Id.Value == roomId)
                    .OrderByDescending(item => item.Score)
                    .Select(item => item.Text.Content?.Trim() ?? "")
                    .Where(content => !string.IsNullOrWhiteSpace(content) && !string.Equals(content, bestMatch.Content?.Trim(), StringComparison.Ordinal))
                    .Distinct()
                    .Take(3)
                    .ToList();
            }

            matches.Add(new RoomTextMatch
            {
                RoomId = room.Id.Value,
                RoomName = room.Name,
                RoomNumber = room.Number,
                MatchedText = bestMatch?.Content ?? "",
                Confidence = confidence,
                Status = status,
                AlternativeTexts = alternatives
            });
        }

        return matches;
    }

    private XYZ? GetRoomCenter(Room room)
    {
        var location = (room.Location as LocationPoint)?.Point;
        if (location != null) return location;

        var bbox = room.get_BoundingBox(null);
        if (bbox == null) return null;
        return new XYZ((bbox.Min.X + bbox.Max.X) / 2, (bbox.Min.Y + bbox.Max.Y) / 2, (bbox.Min.Z + bbox.Max.Z) / 2);
    }

    private string NormalizeForMatch(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var normalized = value.Trim().ToUpperInvariant();
        normalized = Regex.Replace(normalized, @"[\s\-_/\\\[\]\(\)【】]+", "");
        normalized = Regex.Replace(normalized, @"[^\p{L}\p{Nd}]+", "");
        return normalized;
    }

    private double ComputeTextSimilarity(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            return 0.0;

        if (left == right) return 1.0;
        if (left.Contains(right) || right.Contains(left)) return 0.9;

        var distance = LevenshteinDistance(left, right);
        var maxLen = Math.Max(left.Length, right.Length);
        if (maxLen == 0) return 0.0;
        return Math.Max(0.0, 1.0 - (double)distance / maxLen);
    }

    private int LevenshteinDistance(string a, string b)
    {
        var d = new int[a.Length + 1, b.Length + 1];
        for (int i = 0; i <= a.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= b.Length; j++) d[0, j] = j;

        for (int i = 1; i <= a.Length; i++)
        {
            for (int j = 1; j <= b.Length; j++)
            {
                int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }

        return d[a.Length, b.Length];
    }

    /// <summary>
    /// 匹配文字到闭合区域（用于自动创建房间）
    /// </summary>
    public List<RegionTextMatch> MatchTextsToRegions(
        List<DwgTextInfo> texts,
        List<ClosedRegion> regions,
        double maxDistance = 10.0)
    {
        var matches = new List<RegionTextMatch>();

        foreach (var region in regions)
        {
            // 查找区域内的文字
            var matchedTexts = new List<(DwgTextInfo text, double distance)>();

            foreach (var text in texts)
            {
                var dx = text.X - region.CenterX;
                var dy = text.Y - region.CenterY;
                var distance = Math.Sqrt(dx * dx + dy * dy);

                if (distance < maxDistance)
                {
                    matchedTexts.Add((text, distance));
                }
            }

            // 按距离排序
            matchedTexts.Sort((a, b) => a.distance.CompareTo(b.distance));

            var bestMatch = matchedTexts.Count > 0 ? matchedTexts[0].text : null;

            matches.Add(new RegionTextMatch
            {
                Region = region,
                MatchedText = bestMatch?.Content ?? "",
                Confidence = bestMatch != null 
                    ? Math.Max(0, 1.0 - matchedTexts[0].distance / maxDistance) 
                    : 0.0,
                Status = bestMatch != null ? MatchStatus.Matched : MatchStatus.NoText
            });
        }

        return matches;
    }

    /// <summary>
    /// 判断点是否在房间内
    /// </summary>
    private bool IsPointInRoom(XYZ point, Room room)
    {
        try
        {
            // 优先使用 Revit API 精确判断
            if (room.IsPointInRoom(point))
                return true;
            
            // fallback: 用 BoundingBox 粗略判断（XY 平面）
            var bbox = room.get_BoundingBox(null);
            if (bbox != null)
            {
                return point.X >= bbox.Min.X && point.X <= bbox.Max.X &&
                       point.Y >= bbox.Min.Y && point.Y <= bbox.Max.Y;
            }
            
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 批量更新房间名称
    /// </summary>
    public Dictionary<long, bool> UpdateRoomNames(
        Document document, 
        List<RoomTextMatch> matches)
    {
        var results = new Dictionary<long, bool>();

        using var transaction = new Transaction(document, "批量更新房间名称");
        transaction.Start();

        try
        {
            foreach (var match in matches)
            {
                if (match.Status != MatchStatus.Matched && match.Status != MatchStatus.MultipleTexts)
                    continue;

                var elementId = new ElementId(match.RoomId);
                var room = document.GetElement(elementId) as Room;

                if (room != null)
                {
                    room.Name = match.MatchedText;
                    results[match.RoomId] = true;
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
}

/// <summary>
/// DWG 文字信息
/// </summary>
public class DwgTextInfo
{
    public string Content { get; set; } = string.Empty;
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
    public double Height { get; set; }
    public double Rotation { get; set; }
    public string LayerName { get; set; } = string.Empty;
}

/// <summary>
/// DWG 曲线信息
/// </summary>
public class DwgCurveInfo
{
    public CurveType Type { get; set; }
    public double StartX { get; set; }
    public double StartY { get; set; }
    public double EndX { get; set; }
    public double EndY { get; set; }
    public double CenterX { get; set; }
    public double CenterY { get; set; }
    public double Radius { get; set; }
    public string LayerName { get; set; } = string.Empty;
}

/// <summary>
/// 曲线类型
/// </summary>
public enum CurveType
{
    Line,
    Arc,
    Circle
}

/// <summary>
/// 闭合区域
/// </summary>
public class ClosedRegion
{
    public List<DwgCurveInfo> Curves { get; set; } = new();
    public double CenterX { get; set; }
    public double CenterY { get; set; }
    public double Area { get; set; }
}

/// <summary>
/// Revit 坐标系中的文字
/// </summary>
public class TextInRevit
{
    public string Content { get; set; } = string.Empty;
    public XYZ Position { get; set; } = XYZ.Zero;
    public double Height { get; set; }
    public string LayerName { get; set; } = string.Empty;
}

/// <summary>
/// 房间-文字匹配结果
/// </summary>
public class RoomTextMatch
{
    public long RoomId { get; set; }
    public string RoomName { get; set; } = string.Empty;
    public string RoomNumber { get; set; } = string.Empty;
    public string MatchedText { get; set; } = string.Empty;
    public double Confidence { get; set; }
    public MatchStatus Status { get; set; }
    public List<string> AlternativeTexts { get; set; } = new();
}

internal class RoomTextCandidate
{
    public Room Room { get; set; } = null!;
    public TextInRevit Text { get; set; } = null!;
    public double Distance { get; set; }
    public bool IsInsideRoom { get; set; }
    public double Score { get; set; }
}

internal static class PresetKeywordMatcher
{
    private static readonly Dictionary<string, string[]> KeywordGroups = new()
    {
        ["OFFICE"] = new[] { "办公室", "办公", "OFFICE", "OFF" },
        ["MEETING"] = new[] { "会议室", "会议", "MEETING", "MEET" },
        ["RESTROOM"] = new[] { "卫生间", "洗手间", "厕所", "RESTROOM", "TOILET", "WC" },
        ["CORRIDOR"] = new[] { "走廊", "过道", "CORRIDOR", "HALL" },
        ["STAIR"] = new[] { "楼梯", "楼梯间", "STAIR", "STAIRCASE" },
        ["STORAGE"] = new[] { "仓库", "储藏", "STORAGE", "STORE" },
        ["EQUIPMENT"] = new[] { "设备", "机房", "ELECTRIC", "MEP", "机电" }
    };

    public static string GetKeywordGroup(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        var upper = text.ToUpperInvariant();
        foreach (var kv in KeywordGroups)
        {
            if (kv.Value.Any(keyword => upper.Contains(keyword.ToUpperInvariant())))
                return kv.Key;
        }
        return string.Empty;
    }
}

/// <summary>
/// 区域-文字匹配结果
/// </summary>
public class RegionTextMatch
{
    public ClosedRegion Region { get; set; } = new();
    public string MatchedText { get; set; } = string.Empty;
    public double Confidence { get; set; }
    public MatchStatus Status { get; set; }
}

/// <summary>
/// 匹配状态
/// </summary>
public enum MatchStatus
{
    Matched,        // 已匹配
    NoText,         // 未找到文字
    MultipleTexts,  // 多个文字
    LowConfidence   // 低置信度
}
