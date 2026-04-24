using OfficeOpenXml;
using OfficeOpenXml.Style;
using RoomManager.Models;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace RoomManager.Services;

/// <summary>
/// Excel 导入导出服务
/// 使用 EPPlus 库处理 Excel 文件
/// </summary>
public class ExcelService
{
    public ExcelService()
    {
        // EPPlus 8 许可证设置（非商业用途）
        ExcelPackage.License.SetNonCommercialPersonal("RoomManager User");
    }

    /// <summary>
    /// 导出房间数据到 Excel
    /// </summary>
    /// <param name="rooms">房间列表</param>
    /// <param name="filePath">导出文件路径</param>
    /// <param name="includeCustomParams">是否包含自定义参数</param>
    public void ExportToExcel(IEnumerable<RoomData> rooms, string filePath, bool includeCustomParams = true)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("房间列表");

        var roomList = rooms.ToList();
        var customParamNames = includeCustomParams
            ? roomList.SelectMany(r => r.CustomParameters.Keys).Distinct().OrderBy(k => k).ToList()
            : new List<string>();

        // 设置表头
        var headers = new List<string>
        {
            "Element ID", "房间名称", "房间编号", "楼层", "面积(m²)", "体积(m³)", "阶段", "分类", "是否完整"
        };

        if (includeCustomParams)
        {
            headers.AddRange(customParamNames);
        }

        // 写入表头
        for (int i = 0; i < headers.Count; i++)
        {
            var cell = worksheet.Cells[1, i + 1];
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
            cell.Style.Font.Color.SetColor(Color.White);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        // 写入数据
        for (int row = 0; row < roomList.Count; row++)
        {
            var room = roomList[row];
            var dataRow = row + 2;

            worksheet.Cells[dataRow, 1].Value = room.ElementId;
            worksheet.Cells[dataRow, 2].Value = room.Name;
            worksheet.Cells[dataRow, 3].Value = room.Number;
            worksheet.Cells[dataRow, 4].Value = room.Level;
            worksheet.Cells[dataRow, 5].Value = Math.Round(room.Area, 2);
            worksheet.Cells[dataRow, 6].Value = Math.Round(room.Volume, 2);
            worksheet.Cells[dataRow, 7].Value = room.Phase;
            worksheet.Cells[dataRow, 8].Value = RoomCategoryHelper.GetDisplayName(room.Category);
            worksheet.Cells[dataRow, 9].Value = room.IsComplete ? "是" : "否";

            // 自定义参数
            if (includeCustomParams)
            {
                for (int i = 0; i < customParamNames.Count; i++)
                {
                    var paramName = customParamNames[i];
                    if (room.CustomParameters.TryGetValue(paramName, out var value))
                    {
                        worksheet.Cells[dataRow, 10 + i].Value = value?.ToString() ?? "";
                    }
                }
            }
        }

        // 自动调整列宽
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

        // 设置表格样式
        var tableRange = worksheet.Cells[1, 1, roomList.Count + 1, headers.Count];
        var table = worksheet.Tables.Add(tableRange, "RoomTable");
        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

        // 条件格式：不完整的房间标红
        var incompleteRange = worksheet.Cells[2, 9, roomList.Count + 1, 9];
        var conditionalFormat = incompleteRange.ConditionalFormatting.AddEqual();
        conditionalFormat.Formula = "\"否\"";
        conditionalFormat.Style.Font.Color.SetColor(System.Drawing.Color.Red);

        // 添加统计工作表
        AddStatisticsSheet(package, roomList);

        // 添加分类统计工作表
        AddCategorySheet(package, roomList);

        // 保存文件
        package.SaveAs(new FileInfo(filePath));
    }

    /// <summary>
    /// 从 Excel 导入房间数据
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>导入的房间更新数据</returns>
    public List<RoomUpdateData> ImportFromExcel(string filePath)
    {
        var updates = new List<RoomUpdateData>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();

        if (worksheet == null || worksheet.Dimension == null)
            return updates;

        // 读取表头
        var headerCount = worksheet.Dimension.End.Column;
        var headers = new Dictionary<string, int>();
        for (int col = 1; col <= headerCount; col++)
        {
            headers[worksheet.Cells[1, col].Text] = col;
        }

        // 验证必要列
        if (!headers.ContainsKey("Element ID"))
            throw new InvalidDataException("Excel 文件缺少 'Element ID' 列");

        // 读取数据行
        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            var elementIdText = worksheet.Cells[row, headers["Element ID"]].Text;
            if (string.IsNullOrWhiteSpace(elementIdText))
                continue;

            if (!long.TryParse(elementIdText, out var elementId))
                continue;

            var update = new RoomUpdateData
            {
                ElementId = elementId,
                Name = headers.ContainsKey("房间名称") ? worksheet.Cells[row, headers["房间名称"]].Text : null,
                Number = headers.ContainsKey("房间编号") ? worksheet.Cells[row, headers["房间编号"]].Text : null
            };

            // 读取自定义参数
            var customParamColumns = headers
                .Where(h => !new[] { "Element ID", "房间名称", "房间编号", "楼层", "面积(m²)", "体积(m³)", "阶段", "分类", "是否完整" }.Contains(h.Key))
                .ToList();

            foreach (var (paramName, colIndex) in customParamColumns)
            {
                var value = worksheet.Cells[row, colIndex].Text;
                if (!string.IsNullOrWhiteSpace(value))
                {
                    update.CustomParameters[paramName] = value;
                }
            }

            updates.Add(update);
        }

        return updates;
    }

    /// <summary>
    /// 导出 Excel 模板（空表，供用户填写）
    /// </summary>
    /// <param name="filePath">模板文件路径</param>
    /// <param name="rooms">房间列表（用于生成行）</param>
    /// <param name="customParamNames">自定义参数名称列表</param>
    public void ExportTemplate(string filePath, IEnumerable<RoomData> rooms, IEnumerable<string>? customParamNames = null)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("房间导入模板");

        var roomList = rooms.ToList();
        var paramList = customParamNames?.ToList() ?? new List<string>();

        // 设置表头
        var headers = new List<string> { "Element ID", "房间名称", "房间编号" };
        headers.AddRange(paramList);

        for (int i = 0; i < headers.Count; i++)
        {
            var cell = worksheet.Cells[1, i + 1];
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
            cell.Style.Font.Color.SetColor(Color.White);
        }

        // 写入房间 ID（只读）
        for (int row = 0; row < roomList.Count; row++)
        {
            worksheet.Cells[row + 2, 1].Value = roomList[row].ElementId;
            worksheet.Cells[row + 2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row + 2, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
        }

        // 添加说明工作表
        var instructionSheet = package.Workbook.Worksheets.Add("使用说明");
        instructionSheet.Cells[1, 1].Value = "房间导入模板使用说明";
        instructionSheet.Cells[1, 1].Style.Font.Bold = true;
        instructionSheet.Cells[1, 1].Style.Font.Size = 14;

        var instructions = new[]
        {
            "",
            "1. Element ID 列为只读，请勿修改",
            "2. 在「房间名称」和「房间编号」列填写或修改房间信息",
            "3. 自定义参数列可根据需要填写",
            "4. 保存文件后，在插件中选择「从 Excel 导入」",
            "5. 插件将根据 Element ID 更新对应的房间信息",
            "",
            "注意事项：",
            "- 请勿删除或移动 Element ID 列",
            "- 请勿修改文件结构",
            "- 导入前请备份原始模型"
        };

        for (int i = 0; i < instructions.Length; i++)
        {
            instructionSheet.Cells[i + 2, 1].Value = instructions[i];
        }

        instructionSheet.Cells[instructionSheet.Dimension.Address].AutoFitColumns();

        package.SaveAs(new FileInfo(filePath));
    }

    /// <summary>
    /// 添加统计工作表
    /// </summary>
    private void AddStatisticsSheet(ExcelPackage package, List<RoomData> rooms)
    {
        var worksheet = package.Workbook.Worksheets.Add("统计概览");

        // 标题
        worksheet.Cells[1, 1].Value = "房间统计概览";
        worksheet.Cells[1, 1].Style.Font.Bold = true;
        worksheet.Cells[1, 1].Style.Font.Size = 16;

        // 基础统计
        var stats = new Dictionary<string, object>
        {
            ["总房间数"] = rooms.Count,
            ["已完整录入"] = rooms.Count(r => r.IsComplete),
            ["待完善"] = rooms.Count(r => !r.IsComplete),
            ["总面积(m²)"] = Math.Round(rooms.Sum(r => r.Area), 2),
            ["总体积(m³)"] = Math.Round(rooms.Sum(r => r.Volume), 2),
            ["平均面积(m²)"] = rooms.Count > 0 ? Math.Round(rooms.Average(r => r.Area), 2) : 0
        };

        worksheet.Cells[3, 1].Value = "基础统计";
        worksheet.Cells[3, 1].Style.Font.Bold = true;

        int row = 4;
        foreach (var (name, value) in stats)
        {
            worksheet.Cells[row, 1].Value = name;
            worksheet.Cells[row, 2].Value = value;
            row++;
        }

        // 楼层统计
        worksheet.Cells[row + 1, 1].Value = "楼层统计";
        worksheet.Cells[row + 1, 1].Style.Font.Bold = true;
        row += 2;

        var levelStats = rooms
            .GroupBy(r => r.Level)
            .OrderBy(g => g.Key)
            .Select(g => new { Level = g.Key, Count = g.Count(), Area = g.Sum(r => r.Area) });

        worksheet.Cells[row, 1].Value = "楼层";
        worksheet.Cells[row, 2].Value = "房间数";
        worksheet.Cells[row, 3].Value = "面积(m²)";
        worksheet.Cells[row, 1].Style.Font.Bold = true;
        worksheet.Cells[row, 2].Style.Font.Bold = true;
        worksheet.Cells[row, 3].Style.Font.Bold = true;
        row++;

        foreach (var stat in levelStats)
        {
            worksheet.Cells[row, 1].Value = stat.Level;
            worksheet.Cells[row, 2].Value = stat.Count;
            worksheet.Cells[row, 3].Value = Math.Round(stat.Area, 2);
            row++;
        }

        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    }

    /// <summary>
    /// 添加分类统计工作表
    /// </summary>
    private void AddCategorySheet(ExcelPackage package, List<RoomData> rooms)
    {
        var worksheet = package.Workbook.Worksheets.Add("分类统计");

        // 标题
        worksheet.Cells[1, 1].Value = "房间分类统计";
        worksheet.Cells[1, 1].Style.Font.Bold = true;
        worksheet.Cells[1, 1].Style.Font.Size = 16;

        // 表头
        worksheet.Cells[3, 1].Value = "分类";
        worksheet.Cells[3, 2].Value = "数量";
        worksheet.Cells[3, 3].Value = "占比";
        worksheet.Cells[3, 4].Value = "总面积(m²)";

        for (int col = 1; col <= 4; col++)
        {
            worksheet.Cells[3, col].Style.Font.Bold = true;
            worksheet.Cells[3, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[3, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
            worksheet.Cells[3, col].Style.Font.Color.SetColor(Color.White);
        }

        // 分类统计
        var categoryStats = rooms
            .GroupBy(r => r.Category)
            .Select(g => new
            {
                Category = g.Key,
                Count = g.Count(),
                Area = g.Sum(r => r.Area)
            })
            .OrderByDescending(g => g.Count);

        int row = 4;
        var total = rooms.Count;
        foreach (var stat in categoryStats)
        {
            worksheet.Cells[row, 1].Value = RoomCategoryHelper.GetDisplayName(stat.Category);
            worksheet.Cells[row, 2].Value = stat.Count;
            worksheet.Cells[row, 3].Value = total > 0 ? $"{Math.Round((double)stat.Count / total * 100, 1)}%" : "0%";
            worksheet.Cells[row, 4].Value = Math.Round(stat.Area, 2);
            row++;
        }

        // 添加饼图
        var chart = worksheet.Drawings.AddChart("分类饼图", OfficeOpenXml.Drawing.Chart.eChartType.Pie);
        chart.SetPosition(row + 1, 0, 0, 0);
        chart.SetSize(400, 300);
        chart.Title.Text = "房间分类分布";

        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    }
}

/// <summary>
/// 房间更新数据（用于导入）
/// </summary>
public class RoomUpdateData
{
    public long ElementId { get; set; }
    public string? Name { get; set; }
    public string? Number { get; set; }
    public Dictionary<string, object?> CustomParameters { get; set; } = new();
}
