using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Font;
using RoomManager.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoomManager.Services;

/// <summary>
/// PDF 报告生成服务
/// 使用 iText 7/9 库
/// </summary>
public class PdfReportService
{
    /// <summary>
    /// 生成房间报告
    /// </summary>
    /// <param name="rooms">房间列表</param>
    /// <param name="filePath">输出文件路径</param>
    /// <param name="title">报告标题</param>
    /// <param name="includeThumbnails">是否包含略缩图</param>
    public void GenerateRoomReport(IEnumerable<RoomData> rooms, string filePath,
        string title = "房间信息报告", bool includeThumbnails = false)
    {
        var roomList = rooms.ToList();

        using var writer = new PdfWriter(filePath);
        using var pdf = new PdfDocument(writer);
        using var document = new Document(pdf);

        // 标题
        var titleParagraph = new Paragraph(title)
            .SetFontSize(24)
            .SimulateBold()
            .SetTextAlignment(TextAlignment.CENTER)
            .SetMarginBottom(20);
        document.Add(titleParagraph);

        // 生成时间
        var timeParagraph = new Paragraph($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            .SetFontSize(10)
            .SetFontColor(iText.Kernel.Colors.ColorConstants.GRAY)
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetMarginBottom(30);
        document.Add(timeParagraph);

        // 统计概览
        AddStatisticsSection(document, roomList);

        // 楼层分组统计
        AddLevelStatisticsSection(document, roomList);

        // 分类统计
        AddCategoryStatisticsSection(document, roomList);

        // 房间明细表
        AddRoomDetailsTable(document, roomList);

        // 新页面：问题房间列表
        document.Add(new AreaBreak());
        AddIncompleteRoomsSection(document, roomList);

        document.Close();
    }

    /// <summary>
    /// 生成校验报告
    /// </summary>
    public void GenerateValidationReport(ValidationReport report, string filePath)
    {
        using var writer = new PdfWriter(filePath);
        using var pdf = new PdfDocument(writer);
        using var document = new Document(pdf);

        // 标题
        document.Add(new Paragraph("数据校验报告")
            .SetFontSize(24)
            .SimulateBold()
            .SetTextAlignment(TextAlignment.CENTER)
            .SetMarginBottom(20));

        // 生成时间
        document.Add(new Paragraph($"生成时间: {report.ValidationTime:yyyy-MM-dd HH:mm:ss}")
            .SetFontSize(10)
            .SetFontColor(iText.Kernel.Colors.ColorConstants.GRAY)
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetMarginBottom(30));

        // 统计概览
        var statsTable = new Table(4).UseAllAvailableWidth();
        statsTable.AddCell(CreateStatCell("总房间数", report.TotalRooms.ToString()));
        statsTable.AddCell(CreateStatCell("校验通过", report.ValidRoomCount.ToString(), iText.Kernel.Colors.ColorConstants.GREEN));
        statsTable.AddCell(CreateStatCell("错误", report.ErrorCount.ToString(), iText.Kernel.Colors.ColorConstants.RED));
        statsTable.AddCell(CreateStatCell("警告", report.WarningCount.ToString(), iText.Kernel.Colors.ColorConstants.ORANGE));
        document.Add(statsTable);

        document.Add(new Paragraph($"通过率: {report.ValidRate:F1}%")
            .SetFontSize(14)
            .SetMarginTop(20)
            .SetMarginBottom(20));

        // 问题详情
        if (report.RoomResults.Count > 0)
        {
            document.Add(new Paragraph("问题详情")
                .SetFontSize(16)
                .SimulateBold()
                .SetMarginTop(30)
                .SetMarginBottom(10));

            var table = new Table(4).UseAllAvailableWidth();
            table.AddHeaderCell(CreateHeaderCell("房间 ID"));
            table.AddHeaderCell(CreateHeaderCell("字段"));
            table.AddHeaderCell(CreateHeaderCell("严重程度"));
            table.AddHeaderCell(CreateHeaderCell("问题描述"));

            foreach (var (roomId, results) in report.RoomResults)
            {
                foreach (var result in results)
                {
                    table.AddCell(roomId.ToString());
                    table.AddCell(result.FieldName);
                    table.AddCell(result.Severity.ToString());
                    table.AddCell(result.Message);
                }
            }

            document.Add(table);
        }

        document.Close();
    }

    /// <summary>
    /// 添加统计概览
    /// </summary>
    private void AddStatisticsSection(Document document, List<RoomData> rooms)
    {
        document.Add(new Paragraph("统计概览")
            .SetFontSize(16)
            .SimulateBold()
            .SetMarginBottom(10));

        var statsTable = new Table(3).UseAllAvailableWidth();
        statsTable.AddCell(CreateStatCell("总房间数", rooms.Count.ToString()));
        statsTable.AddCell(CreateStatCell("已录入", rooms.Count(r => r.IsComplete).ToString()));
        statsTable.AddCell(CreateStatCell("待录入", rooms.Count(r => !r.IsComplete).ToString()));

        document.Add(statsTable);
        document.Add(new Paragraph($"总面积: {rooms.Sum(r => r.Area):F2} m²")
            .SetMarginTop(10)
            .SetMarginBottom(20));
    }

    /// <summary>
    /// 添加楼层统计
    /// </summary>
    private void AddLevelStatisticsSection(Document document, List<RoomData> rooms)
    {
        document.Add(new Paragraph("楼层统计")
            .SetFontSize(16)
            .SimulateBold()
            .SetMarginTop(20)
            .SetMarginBottom(10));

        var table = new Table(3).UseAllAvailableWidth();
        table.AddHeaderCell(CreateHeaderCell("楼层"));
        table.AddHeaderCell(CreateHeaderCell("房间数"));
        table.AddHeaderCell(CreateHeaderCell("面积 (m²)"));

        var levelStats = rooms
            .GroupBy(r => r.Level)
            .OrderBy(g => g.Key);

        foreach (var stat in levelStats)
        {
            table.AddCell(stat.Key);
            table.AddCell(stat.Count().ToString());
            table.AddCell(stat.Sum(r => r.Area).ToString("F2"));
        }

        document.Add(table);
    }

    /// <summary>
    /// 添加分类统计
    /// </summary>
    private void AddCategoryStatisticsSection(Document document, List<RoomData> rooms)
    {
        document.Add(new Paragraph("房间分类统计")
            .SetFontSize(16)
            .SimulateBold()
            .SetMarginTop(20)
            .SetMarginBottom(10));

        var table = new Table(3).UseAllAvailableWidth();
        table.AddHeaderCell(CreateHeaderCell("分类"));
        table.AddHeaderCell(CreateHeaderCell("数量"));
        table.AddHeaderCell(CreateHeaderCell("占比"));

        var categoryStats = rooms
            .GroupBy(r => r.Category)
            .OrderByDescending(g => g.Count());

        var total = rooms.Count;
        foreach (var stat in categoryStats)
        {
            table.AddCell(RoomCategoryHelper.GetDisplayName(stat.Key));
            table.AddCell(stat.Count().ToString());
            table.AddCell(total > 0 ? $"{(double)stat.Count() / total * 100:F1}%" : "0%");
        }

        document.Add(table);
    }

    /// <summary>
    /// 添加房间明细表
    /// </summary>
    private void AddRoomDetailsTable(Document document, List<RoomData> rooms)
    {
        document.Add(new AreaBreak());

        document.Add(new Paragraph("房间明细")
            .SetFontSize(16)
            .SimulateBold()
            .SetMarginBottom(10));

        var table = new Table(6).UseAllAvailableWidth();
        table.AddHeaderCell(CreateHeaderCell("编号"));
        table.AddHeaderCell(CreateHeaderCell("名称"));
        table.AddHeaderCell(CreateHeaderCell("楼层"));
        table.AddHeaderCell(CreateHeaderCell("面积 (m²)"));
        table.AddHeaderCell(CreateHeaderCell("分类"));
        table.AddHeaderCell(CreateHeaderCell("状态"));

        foreach (var room in rooms.OrderBy(r => r.Level).ThenBy(r => r.Number))
        {
            table.AddCell(room.Number);
            table.AddCell(room.Name);
            table.AddCell(room.Level);
            table.AddCell(room.Area.ToString("F2"));
            table.AddCell(RoomCategoryHelper.GetDisplayName(room.Category));
            table.AddCell(room.IsComplete ? "OK" : "待完善");
        }

        document.Add(table);
    }

    /// <summary>
    /// 添加待完善房间列表
    /// </summary>
    private void AddIncompleteRoomsSection(Document document, List<RoomData> rooms)
    {
        var incompleteRooms = rooms.Where(r => !r.IsComplete).ToList();

        if (incompleteRooms.Count == 0)
        {
            document.Add(new Paragraph("所有房间信息已完整录入 OK")
                .SetFontSize(14)
                .SetFontColor(iText.Kernel.Colors.ColorConstants.GREEN)
                .SetTextAlignment(TextAlignment.CENTER));
            return;
        }

        document.Add(new Paragraph($"待完善房间 ({incompleteRooms.Count} 间)")
            .SetFontSize(16)
            .SimulateBold()
            .SetMarginBottom(10));

        var table = new Table(4).UseAllAvailableWidth();
        table.AddHeaderCell(CreateHeaderCell("编号"));
        table.AddHeaderCell(CreateHeaderCell("名称"));
        table.AddHeaderCell(CreateHeaderCell("楼层"));
        table.AddHeaderCell(CreateHeaderCell("问题"));

        foreach (var room in incompleteRooms)
        {
            var issues = new List<string>();
            if (string.IsNullOrWhiteSpace(room.Name)) issues.Add("名称缺失");
            if (string.IsNullOrWhiteSpace(room.Number)) issues.Add("编号缺失");

            table.AddCell(room.Number);
            table.AddCell(room.Name);
            table.AddCell(room.Level);
            table.AddCell(string.Join(", ", issues));
        }

        document.Add(table);
    }

    /// <summary>
    /// 创建统计单元格
    /// </summary>
    private Cell CreateStatCell(string label, string value, iText.Kernel.Colors.Color? color = null)
    {
        var cell = new Cell();
        cell.Add(new Paragraph(label).SetFontSize(10).SetFontColor(iText.Kernel.Colors.ColorConstants.GRAY));
        cell.Add(new Paragraph(value).SetFontSize(18).SimulateBold());
        if (color != null)
        {
            cell.Add(new Paragraph(value).SetFontSize(18).SimulateBold().SetFontColor(color));
        }
        return cell.SetPadding(10);
    }

    /// <summary>
    /// 创建表头单元格
    /// </summary>
    private Cell CreateHeaderCell(string text)
    {
        return new Cell()
            .Add(new Paragraph(text).SimulateBold())
            .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY)
            .SetPadding(5);
    }
}
