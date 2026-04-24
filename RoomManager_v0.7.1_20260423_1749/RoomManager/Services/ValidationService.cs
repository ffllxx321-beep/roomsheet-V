using RoomManager.Models;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace RoomManager.Services;

/// <summary>
/// 数据校验服务
/// </summary>
public class ValidationService
{
    /// <summary>
    /// 校验规则配置
    /// </summary>
    public ValidationConfig Config { get; set; } = new();

    /// <summary>
    /// 校验单个房间
    /// </summary>
    public List<ValidationResult> ValidateRoom(RoomData room)
    {
        var results = new List<ValidationResult>();

        // 1. 房间名称校验
        if (Config.CheckNameRequired && string.IsNullOrWhiteSpace(room.Name))
        {
            results.Add(new ValidationResult
            {
                FieldName = "房间名称",
                IsValid = false,
                Message = "房间名称不能为空",
                Severity = ValidationSeverity.Error
            });
        }
        else if (Config.CheckNamePattern && !string.IsNullOrWhiteSpace(room.Name))
        {
            if (!Regex.IsMatch(room.Name, Config.NamePattern))
            {
                results.Add(new ValidationResult
                {
                    FieldName = "房间名称",
                    IsValid = false,
                    Message = $"房间名称不符合规则: {Config.NamePatternHint}",
                    Severity = ValidationSeverity.Warning
                });
            }
        }

        // 2. 房间编号校验
        if (Config.CheckNumberRequired && string.IsNullOrWhiteSpace(room.Number))
        {
            results.Add(new ValidationResult
            {
                FieldName = "房间编号",
                IsValid = false,
                Message = "房间编号不能为空",
                Severity = ValidationSeverity.Error
            });
        }
        else if (Config.CheckNumberPattern && !string.IsNullOrWhiteSpace(room.Number))
        {
            if (!Regex.IsMatch(room.Number, Config.NumberPattern))
            {
                results.Add(new ValidationResult
                {
                    FieldName = "房间编号",
                    IsValid = false,
                    Message = $"房间编号不符合规则: {Config.NumberPatternHint}",
                    Severity = ValidationSeverity.Warning
                });
            }
        }

        // 3. 面积校验
        if (Config.CheckMinArea && room.Area < Config.MinArea)
        {
            results.Add(new ValidationResult
            {
                FieldName = "面积",
                IsValid = false,
                Message = $"面积过小: {room.Area:F2} m² < {Config.MinArea} m²",
                Severity = ValidationSeverity.Warning
            });
        }

        if (Config.CheckMaxArea && room.Area > Config.MaxArea)
        {
            results.Add(new ValidationResult
            {
                FieldName = "面积",
                IsValid = false,
                Message = $"面积过大: {room.Area:F2} m² > {Config.MaxArea} m²",
                Severity = ValidationSeverity.Warning
            });
        }

        // 4. 楼层校验
        if (Config.CheckLevelRequired && string.IsNullOrWhiteSpace(room.Level))
        {
            results.Add(new ValidationResult
            {
                FieldName = "楼层",
                IsValid = false,
                Message = "楼层信息缺失",
                Severity = ValidationSeverity.Error
            });
        }

        // 5. 自定义参数校验
        foreach (var (paramName, rule) in Config.CustomParamRules)
        {
            if (room.CustomParameters.TryGetValue(paramName, out var value))
            {
                if (rule.Required && (value == null || string.IsNullOrWhiteSpace(value.ToString())))
                {
                    results.Add(new ValidationResult
                    {
                        FieldName = paramName,
                        IsValid = false,
                        Message = $"{paramName} 不能为空",
                        Severity = ValidationSeverity.Error
                    });
                }
                else if (!string.IsNullOrWhiteSpace(value?.ToString()) && !string.IsNullOrWhiteSpace(rule.Pattern))
                {
                    if (!Regex.IsMatch(value.ToString()!, rule.Pattern))
                    {
                        results.Add(new ValidationResult
                        {
                            FieldName = paramName,
                            IsValid = false,
                            Message = $"{paramName} 格式不正确",
                            Severity = ValidationSeverity.Warning
                        });
                    }
                }
            }
            else if (rule.Required)
            {
                results.Add(new ValidationResult
                {
                    FieldName = paramName,
                    IsValid = false,
                    Message = $"{paramName} 缺失",
                    Severity = ValidationSeverity.Error
                });
            }
        }

        return results;
    }

    /// <summary>
    /// 批量校验房间
    /// </summary>
    public ValidationReport ValidateRooms(IEnumerable<RoomData> rooms)
    {
        var roomList = rooms.ToList();
        var report = new ValidationReport
        {
            TotalRooms = roomList.Count,
            ValidationTime = DateTime.Now
        };

        foreach (var room in roomList)
        {
            var results = ValidateRoom(room);
            if (results.Any(r => !r.IsValid))
            {
                report.RoomResults[room.ElementId] = results.Where(r => !r.IsValid).ToList();
            }
        }

        report.ErrorCount = report.RoomResults.Values.Sum(r => r.Count(v => v.Severity == ValidationSeverity.Error));
        report.WarningCount = report.RoomResults.Values.Sum(r => r.Count(v => v.Severity == ValidationSeverity.Warning));
        report.ValidRoomCount = roomList.Count - report.RoomResults.Count;

        return report;
    }

    /// <summary>
    /// 检查房间编号重复
    /// </summary>
    public List<DuplicateResult> CheckDuplicateNumbers(IEnumerable<RoomData> rooms)
    {
        var duplicates = rooms
            .GroupBy(r => r.Number)
            .Where(g => g.Count() > 1 && !string.IsNullOrWhiteSpace(g.Key))
            .Select(g => new DuplicateResult
            {
                Value = g.Key,
                Count = g.Count(),
                RoomIds = g.Select(r => r.ElementId).ToList()
            })
            .ToList();

        return duplicates;
    }

    /// <summary>
    /// 检查房间名称重复
    /// </summary>
    public List<DuplicateResult> CheckDuplicateNames(IEnumerable<RoomData> rooms)
    {
        var duplicates = rooms
            .GroupBy(r => r.Name)
            .Where(g => g.Count() > 1 && !string.IsNullOrWhiteSpace(g.Key))
            .Select(g => new DuplicateResult
            {
                Value = g.Key,
                Count = g.Count(),
                RoomIds = g.Select(r => r.ElementId).ToList()
            })
            .ToList();

        return duplicates;
    }
}

/// <summary>
/// 校验配置
/// </summary>
public class ValidationConfig
{
    // 名称校验
    public bool CheckNameRequired { get; set; } = true;
    public bool CheckNamePattern { get; set; } = false;
    public string NamePattern { get; set; } = @".+";
    public string NamePatternHint { get; set; } = "任意非空字符";

    // 编号校验
    public bool CheckNumberRequired { get; set; } = true;
    public bool CheckNumberPattern { get; set; } = false;
    public string NumberPattern { get; set; } = @"^[A-Z0-9\-]+$";
    public string NumberPatternHint { get; set; } = "大写字母、数字、横线";

    // 面积校验
    public bool CheckMinArea { get; set; } = false;
    public double MinArea { get; set; } = 1.0;
    public bool CheckMaxArea { get; set; } = false;
    public double MaxArea { get; set; } = 1000.0;

    // 楼层校验
    public bool CheckLevelRequired { get; set; } = true;

    // 自定义参数规则
    public Dictionary<string, ParamValidationRule> CustomParamRules { get; set; } = new();
}

/// <summary>
/// 参数校验规则
/// </summary>
public class ParamValidationRule
{
    public bool Required { get; set; }
    public string Pattern { get; set; } = string.Empty;
    public string Hint { get; set; } = string.Empty;
}

/// <summary>
/// 校验结果
/// </summary>
public class ValidationResult
{
    public string FieldName { get; set; } = string.Empty;
    public bool IsValid { get; set; }
    public string Message { get; set; } = string.Empty;
    public ValidationSeverity Severity { get; set; }
}

/// <summary>
/// 校验严重程度
/// </summary>
public enum ValidationSeverity
{
    Info,
    Warning,
    Error
}

/// <summary>
/// 校验报告
/// </summary>
public class ValidationReport
{
    public int TotalRooms { get; set; }
    public int ValidRoomCount { get; set; }
    public int ErrorCount { get; set; }
    public int WarningCount { get; set; }
    public DateTime ValidationTime { get; set; }
    
    public Dictionary<long, List<ValidationResult>> RoomResults { get; set; } = new();

    public double ValidRate => TotalRooms > 0 ? (double)ValidRoomCount / TotalRooms * 100 : 0;

    /// <summary>
    /// 所有问题的扁平列表（用于 UI 绑定）
    /// </summary>
    public List<ValidationIssueItem> AllIssues => RoomResults
        .SelectMany(kv => kv.Value.Select(v => new ValidationIssueItem
        {
            RoomId = kv.Key,
            FieldName = v.FieldName,
            Severity = v.Severity.ToString(),
            Message = v.Message
        }))
        .ToList();
}

/// <summary>
/// 校验问题项（用于 UI 显示）
/// </summary>
public class ValidationIssueItem
{
    public long RoomId { get; set; }
    public string RoomName { get; set; } = string.Empty;
    public string FieldName { get; set; } = string.Empty;
    public string Severity { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
}

/// <summary>
/// 重复检测结果
/// </summary>
public class DuplicateResult
{
    public string Value { get; set; } = string.Empty;
    public int Count { get; set; }
    public List<long> RoomIds { get; set; } = new();
}
