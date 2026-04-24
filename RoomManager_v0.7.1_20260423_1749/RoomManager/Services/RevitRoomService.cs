using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using System.Collections.Generic;
using System.Linq;

namespace RoomManager.Services;

/// <summary>
/// Revit 房间服务
/// </summary>
public class RevitRoomService
{
    private readonly Document _document;

    public RevitRoomService(Document document)
    {
        _document = document;
    }

    /// <summary>
    /// 获取所有房间
    /// </summary>
    public List<RoomData> GetAllRooms()
    {
        var rooms = new List<RoomData>();
        var collector = new FilteredElementCollector(_document)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType();

        foreach (Room room in collector)
        {
            rooms.Add(ConvertToRoomData(room));
        }

        return rooms;
    }

    /// <summary>
    /// 将 Revit Room 转换为 RoomData
    /// </summary>
    public RoomData ConvertToRoomData(Room room)
    {
        var roomData = new RoomData
        {
            ElementId = room.Id.Value,
            Name = room.Name ?? "",
            Number = room.Number ?? "",
            Level = room.Level?.Name ?? "未指定",
            Area = room.Area,
            Volume = room.Volume,
            Phase = GetPhaseName(room),
            Category = RoomCategoryHelper.Classify(room.Name),
            IsComplete = !string.IsNullOrWhiteSpace(room.Name) && !string.IsNullOrWhiteSpace(room.Number)
        };

        // 获取所有参数（包括只读和可编辑）
        try
        {
            foreach (Parameter param in room.Parameters)
            {
                if (param.Definition == null) continue;
                
                var paramName = param.Definition.Name;
                
                // 排除 IFC 参数和阶段化参数
                if (paramName.StartsWith("Ifc", StringComparison.OrdinalIgnoreCase)) continue;
                if (paramName.Contains("Phase") || paramName.Contains("阶段")) continue;
                
                // 获取参数组名
                string groupName;
                try
                {
                    var groupTypeId = param.Definition.GetGroupTypeId();
                    groupName = LabelUtils.GetLabelForGroup(groupTypeId);
                }
                catch
                {
                    groupName = "其他";
                }
                if (string.IsNullOrEmpty(groupName)) groupName = "其他";
                
                // 获取显示值
                var displayValue = GetParameterDisplayValue(param);
                
                roomData.AllParameters.Add(new RoomParameterInfo
                {
                    Name = paramName,
                    DisplayName = paramName,
                    GroupName = groupName,
                    IsReadOnly = param.IsReadOnly,
                    StorageTypeName = param.StorageType.ToString(),
                    DisplayValue = displayValue,
                    OriginalValue = displayValue
                });
                
                // 兼容旧的 CustomParameters
                if (!param.IsReadOnly)
                {
                    roomData.CustomParameters[paramName] = GetParameterValue(param);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"获取参数失败: {ex.Message}");
        }

        return roomData;
    }

    /// <summary>
    /// 安全获取阶段名称
    /// </summary>
    private string GetPhaseName(Room room)
    {
        try
        {
            var phaseId = room.CreatedPhaseId;
            if (phaseId != null && phaseId != ElementId.InvalidElementId)
            {
                var phase = _document.GetElement(phaseId);
                return phase?.Name ?? "";
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"获取阶段失败: {ex.Message}");
        }
        return "";
    }

    /// <summary>
    /// 获取当前视图中的链接 DWG
    /// </summary>
    public List<CADLinkInfo> GetLinkedCADLinks(View view)
    {
        var links = new List<CADLinkInfo>();
        
        // 方式1: 通过 ImportInstance 查找链接的 DWG
        try
        {
            var collector = new FilteredElementCollector(_document)
                .OfClass(typeof(ImportInstance))
                .WhereElementIsNotElementType();

            foreach (ImportInstance import in collector)
            {
                // 只要是链接的就行，不用检查 Category 名称（中文版可能不同）
                if (!import.IsLinked) continue;
                
                try
                {
                    var typeId = import.GetTypeId();
                    var typeElement = _document.GetElement(typeId);
                    
                    string name = typeElement?.Name ?? $"CAD Link {import.Id.Value}";
                    string filePath = "";
                    
                    // 尝试获取文件路径
                    if (typeElement is CADLinkType cadType)
                    {
                        try
                        {
                            var extRef = cadType.GetExternalFileReference();
                            if (extRef != null)
                                filePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetAbsolutePath());
                        }
                        catch { }
                    }
                    
                    if (string.IsNullOrEmpty(filePath))
                        filePath = name; // fallback
                    
                    links.Add(new CADLinkInfo
                    {
                        ElementId = import.Id.Value,
                        Name = name,
                        FilePath = filePath,
                        Transform = import.GetTransform()
                    });
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"处理 CAD 链接 {import.Id.Value} 失败: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"GetLinkedCADLinks 方式1 失败: {ex.Message}");
        }
        
        // 方式2: 如果方式1没找到，尝试通过 RevitLinkInstance 查找（有些 DWG 可能通过这个方式链接）
        if (links.Count == 0)
        {
            try
            {
                var collector2 = new FilteredElementCollector(_document)
                    .OfClass(typeof(ImportInstance));
                
                foreach (ImportInstance import in collector2)
                {
                    try
                    {
                        var typeId = import.GetTypeId();
                        var typeElement = _document.GetElement(typeId);
                        string name = typeElement?.Name ?? $"Import {import.Id.Value}";
                        
                        links.Add(new CADLinkInfo
                        {
                            ElementId = import.Id.Value,
                            Name = name,
                            FilePath = name,
                            Transform = import.GetTransform()
                        });
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetLinkedCADLinks 方式2 失败: {ex.Message}");
            }
        }

        return links;
    }

    /// <summary>
    /// 在当前标高自动放置房间
    /// </summary>
    public List<RoomCreationResult> AutoPlaceRoomsOnLevel(Level level, View view)
    {
        var results = new List<RoomCreationResult>();

        using var transaction = new Transaction(_document, "自动放置房间");
        transaction.Start();

        try
        {
            // 获取当前标高上所有闭合区域
            var phases = _document.Phases;
            var phase = phases.get_Item(phases.Size - 1); // 使用最后一个阶段

            // 使用 Revit API 的自动放置房间功能
            // NewRooms2 返回 ElementId 集合，需要通过 GetElement 获取 Element
            var roomElementIds = _document.Create.NewRooms2(level, phase);

            foreach (ElementId elementId in roomElementIds)
            {
                if (_document.GetElement(elementId) is Room room)
                {
                    results.Add(new RoomCreationResult
                    {
                        Success = true,
                        RoomId = room.Id.Value,
                        Message = $"已创建房间: {room.Name}"
                    });
                }
            }

            transaction.Commit();
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            results.Add(new RoomCreationResult
            {
                Success = false,
                Message = $"创建失败: {ex.Message}"
            });
        }

        return results;
    }

    /// <summary>
    /// 在选定区域放置房间
    /// </summary>
    public List<RoomCreationResult> PlaceRoomsInSelection(BoundingBoxXYZ selectionBox, Level level, View view)
    {
        var results = new List<RoomCreationResult>();

        // 将 BoundingBoxXYZ 转换为 Outline
        var outline = new Outline(selectionBox.Min, selectionBox.Max);

        // 收集选择区域内的墙和房间分隔线
        var walls = new FilteredElementCollector(_document, view.Id)
            .OfCategory(BuiltInCategory.OST_Walls)
            .WherePasses(new BoundingBoxIntersectsFilter(outline))
            .Cast<Wall>()
            .ToList();

        var roomSeparators = new FilteredElementCollector(_document, view.Id)
            .OfCategory(BuiltInCategory.OST_RoomSeparationLines)
            .WherePasses(new BoundingBoxIntersectsFilter(outline))
            .Cast<CurveElement>()
            .ToList();

        // 使用 Revit 内置的 NewRooms2 在选择区域内自动放置房间
        using var transaction = new Transaction(_document, "在选区内放置房间");
        transaction.Start();

        try
        {
            var phases = _document.Phases;
            var phase = phases.get_Item(phases.Size - 1);

            // 用 NewRooms2 自动在所有闭合区域放置房间
            var roomElementIds = _document.Create.NewRooms2(level, phase);

            foreach (ElementId elementId in roomElementIds)
            {
                if (_document.GetElement(elementId) is Room room)
                {
                    // 检查房间是否在选择区域内
                    var roomLocation = room.Location as LocationPoint;
                    if (roomLocation != null)
                    {
                        var pt = roomLocation.Point;
                        if (pt.X >= outline.MinimumPoint.X && pt.X <= outline.MaximumPoint.X &&
                            pt.Y >= outline.MinimumPoint.Y && pt.Y <= outline.MaximumPoint.Y)
                        {
                            results.Add(new RoomCreationResult
                            {
                                Success = true,
                                RoomId = room.Id.Value,
                                Message = $"已创建房间: {room.Name}"
                            });
                        }
                        else
                        {
                            // 不在选区内的房间删除
                            _document.Delete(elementId);
                        }
                    }
                }
            }

            transaction.Commit();
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            results.Add(new RoomCreationResult
            {
                Success = false,
                Message = $"放置房间失败: {ex.Message}"
            });
        }

        return results;
    }

    /// <summary>
    /// 更新房间名称
    /// </summary>
    public bool UpdateRoomName(long roomId, string name)
    {
        var elementId = new ElementId(roomId);
        var room = _document.GetElement(elementId) as Room;
        
        if (room == null) return false;

        using var transaction = new Transaction(_document, "更新房间名称");
        transaction.Start();
        
        try
        {
            room.Name = name;
            transaction.Commit();
            return true;
        }
        catch
        {
            transaction.RollBack();
            return false;
        }
    }

    /// <summary>
    /// 更新房间编号
    /// </summary>
    public bool UpdateRoomNumber(long roomId, string number)
    {
        var elementId = new ElementId(roomId);
        var room = _document.GetElement(elementId) as Room;
        
        if (room == null) return false;

        using var transaction = new Transaction(_document, "更新房间编号");
        transaction.Start();
        
        try
        {
            room.Number = number;
            transaction.Commit();
            return true;
        }
        catch
        {
            transaction.RollBack();
            return false;
        }
    }

    /// <summary>
    /// 定位房间到视图
    /// </summary>
    public void LocateRoomInView(long roomId, View view)
    {
        var elementId = new ElementId(roomId);
        var room = _document.GetElement(elementId) as Room;

        if (room == null) return;

        // 获取房间的 BoundingBox，计算视图缩放范围
        var bbox = room.get_BoundingBox(view);
        if (bbox == null) return;

        // 扩大范围让房间居中显示，留出边距
        var margin = 3.0; // 约 1 米的边距
        var min = new XYZ(bbox.Min.X - margin, bbox.Min.Y - margin, bbox.Min.Z);
        var max = new XYZ(bbox.Max.X + margin, bbox.Max.Y + margin, bbox.Max.Z);

        // 通过修改视图的裁剪区域来定位（不需要 UIDocument）
        try
        {
            using var transaction = new Transaction(_document, "定位房间");
            transaction.Start();

            // 创建视图裁剪框
            var cropBox = new BoundingBoxXYZ
            {
                Min = min,
                Max = max
            };

            // 设置视图裁剪
            view.CropBoxActive = true;
            view.CropBox = cropBox;
            view.CropBoxVisible = true;

            transaction.Commit();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"定位房间失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 获取参数的可读显示值
    /// </summary>
    private string GetParameterDisplayValue(Parameter param)
    {
        if (!param.HasValue) return "";
        
        try
        {
            return param.StorageType switch
            {
                StorageType.String => param.AsString() ?? "",
                StorageType.Integer => param.AsInteger().ToString(),
                StorageType.Double => param.AsValueString() ?? param.AsDouble().ToString("F4"),
                StorageType.ElementId => _document.GetElement(param.AsElementId())?.Name ?? "",
                _ => ""
            };
        }
        catch
        {
            return "";
        }
    }

    private object? GetParameterValue(Parameter param)
    {
        return param.StorageType switch
        {
            StorageType.String => param.AsString(),
            StorageType.Integer => param.AsInteger(),
            StorageType.Double => param.AsDouble(),
            StorageType.ElementId => param.AsElementId().Value,
            _ => null
        };
    }
}

/// <summary>
/// CAD 链接信息
/// </summary>
public class CADLinkInfo
{
    public long ElementId { get; set; }
    public string Name { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public Transform Transform { get; set; } = Transform.Identity;
}

/// <summary>
/// 房间创建结果
/// </summary>
public class RoomCreationResult
{
    public bool Success { get; set; }
    public long RoomId { get; set; }
    public string Message { get; set; } = string.Empty;
}
