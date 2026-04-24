using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace RoomManager.Services;

/// <summary>
/// 批量操作服务
/// </summary>
public class BatchOperationService
{
    private readonly Document _document;

    public BatchOperationService(Document document)
    {
        _document = document;
    }

    /// <summary>
    /// 批量修改房间名称（正则替换）
    /// </summary>
    /// <param name="roomIds">房间 ID 列表</param>
    /// <param name="pattern">正则表达式</param>
    /// <param name="replacement">替换文本</param>
    /// <returns>修改结果</returns>
    public BatchResult BatchRenameByRegex(IEnumerable<long> roomIds, string pattern, string replacement)
    {
        var result = new BatchResult();

        using var transaction = new Transaction(_document, "批量重命名");
        transaction.Start();

        try
        {
            var regex = new Regex(pattern, RegexOptions.IgnoreCase);
            int successCount = 0;
            int failCount = 0;

            foreach (var roomId in roomIds)
            {
                try
                {
                    var elementId = new ElementId(roomId);
                    var room = _document.GetElement(elementId) as Room;

                    if (room == null)
                    {
                        failCount++;
                        result.FailedRooms.Add(new FailedRoomInfo
                        {
                            RoomId = roomId,
                            Reason = "房间不存在"
                        });
                        continue;
                    }

                    var oldName = room.Name;
                    var newName = regex.Replace(oldName, replacement);

                    if (oldName != newName)
                    {
                        room.Name = newName;
                        result.ChangedRooms.Add(new ChangedRoomInfo
                        {
                            RoomId = roomId,
                            OldValue = oldName,
                            NewValue = newName
                        });
                    }

                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    result.FailedRooms.Add(new FailedRoomInfo
                    {
                        RoomId = roomId,
                        Reason = ex.Message
                    });
                }
            }

            transaction.Commit();
            result.SuccessCount = successCount;
            result.FailCount = failCount;
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// 批量修改房间名称（前缀/后缀）
    /// </summary>
    public BatchResult BatchRenamePrefixSuffix(IEnumerable<long> roomIds, string prefix, string suffix, bool removeExisting = false)
    {
        var result = new BatchResult();

        using var transaction = new Transaction(_document, "批量添加前缀后缀");
        transaction.Start();

        try
        {
            int successCount = 0;
            int failCount = 0;

            foreach (var roomId in roomIds)
            {
                try
                {
                    var elementId = new ElementId(roomId);
                    var room = _document.GetElement(elementId) as Room;

                    if (room == null)
                    {
                        failCount++;
                        continue;
                    }

                    var oldName = room.Name;
                    var newName = oldName;

                    // 移除现有前缀/后缀
                    if (removeExisting)
                    {
                        if (!string.IsNullOrEmpty(prefix) && newName.StartsWith(prefix))
                            newName = newName.Substring(prefix.Length);
                        if (!string.IsNullOrEmpty(suffix) && newName.EndsWith(suffix))
                            newName = newName.Substring(0, newName.Length - suffix.Length);
                    }

                    // 添加新前缀/后缀
                    if (!string.IsNullOrEmpty(prefix))
                        newName = prefix + newName;
                    if (!string.IsNullOrEmpty(suffix))
                        newName = newName + suffix;

                    if (oldName != newName)
                    {
                        room.Name = newName;
                        result.ChangedRooms.Add(new ChangedRoomInfo
                        {
                            RoomId = roomId,
                            OldValue = oldName,
                            NewValue = newName
                        });
                    }

                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    result.FailedRooms.Add(new FailedRoomInfo
                    {
                        RoomId = roomId,
                        Reason = ex.Message
                    });
                }
            }

            transaction.Commit();
            result.SuccessCount = successCount;
            result.FailCount = failCount;
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// 批量编号（自动递增）
    /// </summary>
    /// <param name="roomIds">房间 ID 列表（按顺序编号）</param>
    /// <param name="template">编号模板（如 "F{floor}-{index:00}"）</param>
    /// <param name="startIndex">起始编号</param>
    /// <param name="step">步长</param>
    public BatchResult BatchNumbering(IEnumerable<long> roomIds, string template, int startIndex = 1, int step = 1)
    {
        var result = new BatchResult();

        using var transaction = new Transaction(_document, "批量编号");
        transaction.Start();

        try
        {
            int currentIndex = startIndex;
            int successCount = 0;
            int failCount = 0;

            foreach (var roomId in roomIds)
            {
                try
                {
                    var elementId = new ElementId(roomId);
                    var room = _document.GetElement(elementId) as Room;

                    if (room == null)
                    {
                        failCount++;
                        continue;
                    }

                    var oldNumber = room.Number;
                    var floorName = room.Level?.Name ?? "";
                    var floorNumber = ExtractFloorNumber(floorName);

                    // 应用模板
                    var newNumber = template
                        .Replace("{floor}", floorNumber)
                        .Replace("{index}", currentIndex.ToString())
                        .Replace("{index:00}", currentIndex.ToString("00"))
                        .Replace("{index:000}", currentIndex.ToString("000"))
                        .Replace("{name}", room.Name);

                    if (oldNumber != newNumber)
                    {
                        room.Number = newNumber;
                        result.ChangedRooms.Add(new ChangedRoomInfo
                        {
                            RoomId = roomId,
                            OldValue = oldNumber,
                            NewValue = newNumber
                        });
                    }

                    currentIndex += step;
                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    result.FailedRooms.Add(new FailedRoomInfo
                    {
                        RoomId = roomId,
                        Reason = ex.Message
                    });
                }
            }

            transaction.Commit();
            result.SuccessCount = successCount;
            result.FailCount = failCount;
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// 批量设置参数值
    /// </summary>
    public BatchResult BatchSetParameter(IEnumerable<long> roomIds, string parameterName, object value)
    {
        var result = new BatchResult();

        using var transaction = new Transaction(_document, "批量设置参数");
        transaction.Start();

        try
        {
            int successCount = 0;
            int failCount = 0;

            foreach (var roomId in roomIds)
            {
                try
                {
                    var elementId = new ElementId(roomId);
                    var room = _document.GetElement(elementId) as Room;

                    if (room == null)
                    {
                        failCount++;
                        continue;
                    }

                    var param = room.LookupParameter(parameterName);
                    if (param == null)
                    {
                        failCount++;
                        result.FailedRooms.Add(new FailedRoomInfo
                        {
                            RoomId = roomId,
                            Reason = $"参数 '{parameterName}' 不存在"
                        });
                        continue;
                    }

                    if (param.IsReadOnly)
                    {
                        failCount++;
                        result.FailedRooms.Add(new FailedRoomInfo
                        {
                            RoomId = roomId,
                            Reason = $"参数 '{parameterName}' 只读"
                        });
                        continue;
                    }

                    // 根据参数类型设置值
                    var oldValue = GetParameterValue(param);
                    SetParameterValue(param, value);

                    result.ChangedRooms.Add(new ChangedRoomInfo
                    {
                        RoomId = roomId,
                        ParameterName = parameterName,
                        OldValue = oldValue?.ToString() ?? "",
                        NewValue = value?.ToString() ?? ""
                    });

                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    result.FailedRooms.Add(new FailedRoomInfo
                    {
                        RoomId = roomId,
                        Reason = ex.Message
                    });
                }
            }

            transaction.Commit();
            result.SuccessCount = successCount;
            result.FailCount = failCount;
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// 批量删除房间
    /// </summary>
    public BatchResult BatchDeleteRooms(IEnumerable<long> roomIds)
    {
        var result = new BatchResult();

        using var transaction = new Transaction(_document, "批量删除房间");
        transaction.Start();

        try
        {
            int successCount = 0;
            int failCount = 0;

            foreach (var roomId in roomIds)
            {
                try
                {
                    var elementId = new ElementId(roomId);
                    var room = _document.GetElement(elementId) as Room;

                    if (room == null)
                    {
                        failCount++;
                        continue;
                    }

                    _document.Delete(elementId);
                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    result.FailedRooms.Add(new FailedRoomInfo
                    {
                        RoomId = roomId,
                        Reason = ex.Message
                    });
                }
            }

            transaction.Commit();
            result.SuccessCount = successCount;
            result.FailCount = failCount;
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// 批量复制房间参数
    /// </summary>
    public BatchResult BatchCopyParameter(IEnumerable<long> sourceRoomIds, string sourceParam, 
        IEnumerable<long> targetRoomIds, string targetParam)
    {
        var result = new BatchResult();
        var sourceIds = sourceRoomIds.ToList();
        var targetIds = targetRoomIds.ToList();

        if (sourceIds.Count != targetIds.Count)
        {
            result.ErrorMessage = "源房间数量与目标房间数量不匹配";
            return result;
        }

        using var transaction = new Transaction(_document, "批量复制参数");
        transaction.Start();

        try
        {
            int successCount = 0;
            int failCount = 0;

            for (int i = 0; i < sourceIds.Count; i++)
            {
                try
                {
                    var sourceRoom = _document.GetElement(new ElementId(sourceIds[i])) as Room;
                    var targetRoom = _document.GetElement(new ElementId(targetIds[i])) as Room;

                    if (sourceRoom == null || targetRoom == null)
                    {
                        failCount++;
                        continue;
                    }

                    var sourceParamObj = sourceRoom.LookupParameter(sourceParam);
                    var targetParamObj = targetRoom.LookupParameter(targetParam);

                    if (sourceParamObj == null || targetParamObj == null)
                    {
                        failCount++;
                        continue;
                    }

                    var value = GetParameterValue(sourceParamObj);
                    SetParameterValue(targetParamObj, value);

                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    result.FailedRooms.Add(new FailedRoomInfo
                    {
                        RoomId = targetIds[i],
                        Reason = ex.Message
                    });
                }
            }

            transaction.Commit();
            result.SuccessCount = successCount;
            result.FailCount = failCount;
        }
        catch (Exception ex)
        {
            transaction.RollBack();
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// 从楼层名称提取楼层编号
    /// </summary>
    private string ExtractFloorNumber(string levelName)
    {
        var match = Regex.Match(levelName, @"\d+");
        return match.Success ? match.Value : "0";
    }

    /// <summary>
    /// 获取参数值
    /// </summary>
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

    /// <summary>
    /// 设置参数值
    /// </summary>
    private void SetParameterValue(Parameter param, object? value)
    {
        if (value == null) return;

        switch (param.StorageType)
        {
            case StorageType.String:
                param.Set(value.ToString());
                break;
            case StorageType.Integer:
                if (value is int intValue)
                    param.Set(intValue);
                else if (int.TryParse(value.ToString(), out var parsedInt))
                    param.Set(parsedInt);
                break;
            case StorageType.Double:
                if (value is double doubleValue)
                    param.Set(doubleValue);
                else if (double.TryParse(value.ToString(), out var parsedDouble))
                    param.Set(parsedDouble);
                break;
            case StorageType.ElementId:
                if (value is long longValue)
                    param.Set(new ElementId(longValue));
                else if (value is int intIdValue)
                    param.Set(new ElementId(intIdValue));
                break;
        }
    }
}

/// <summary>
/// 批量操作结果
/// </summary>
public class BatchResult
{
    public int SuccessCount { get; set; }
    public int FailCount { get; set; }
    public string? ErrorMessage { get; set; }
    public List<ChangedRoomInfo> ChangedRooms { get; set; } = new();
    public List<FailedRoomInfo> FailedRooms { get; set; } = new();
    public bool HasError => !string.IsNullOrEmpty(ErrorMessage) || FailCount > 0;
}

/// <summary>
/// 已修改房间信息
/// </summary>
public class ChangedRoomInfo
{
    public long RoomId { get; set; }
    public string ParameterName { get; set; } = "Name";
    public string OldValue { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;
}

/// <summary>
/// 失败房间信息
/// </summary>
public class FailedRoomInfo
{
    public long RoomId { get; set; }
    public string Reason { get; set; } = string.Empty;
}
