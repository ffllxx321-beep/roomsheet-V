using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace RoomManager.Services;

/// <summary>
/// 房间略缩图生成服务
/// </summary>
public class ThumbnailService
{
    private readonly Document _document;

    public ThumbnailService(Document document)
    {
        _document = document;
    }

    /// <summary>
    /// 生成房间略缩图
    /// </summary>
    /// <param name="roomId">房间 ElementId</param>
    /// <param name="width">图像宽度</param>
    /// <param name="height">图像高度</param>
    /// <returns>BitmapImage 或 null</returns>
    public BitmapImage? GenerateRoomThumbnail(long roomId, int width = 300, int height = 200)
    {
        var elementId = new ElementId(roomId);
        var room = _document.GetElement(elementId) as Room;

        if (room == null) return null;

        try
        {
            // 获取房间边界
            var boundarySegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
            if (boundarySegments == null || boundarySegments.Count == 0)
                return null;

            // 计算边界范围
            var boundingBox = CalculateBoundingBox(boundarySegments);
            if (boundingBox == null) return null;

            // 创建 Drawing Visual
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                // 背景
                dc.DrawRectangle(Brushes.White, null, new System.Windows.Rect(0, 0, width, height));

                // 绘制房间边界
                var pen = new Pen(Brushes.Black, 2);
                var fillBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 247, 255));

                foreach (var loop in boundarySegments)
                {
                    var points = ConvertToScreenPoints(loop, boundingBox, width, height);
                    if (points.Count < 3) continue;

                    // 填充
                    var geometry = new StreamGeometry();
                    using (var ctx = geometry.Open())
                    {
                        ctx.BeginFigure(points[0], true, true);
                        for (int i = 1; i < points.Count; i++)
                        {
                            ctx.LineTo(points[i], true, false);
                        }
                    }
                    geometry.Freeze();
                    dc.DrawGeometry(fillBrush, pen, geometry);
                }

                // 绘制房间名称和编号
                var centerPoint = CalculateCenter(boundarySegments, boundingBox, width, height);
                var text = $"{room.Name}\n{room.Number}";
                var formattedText = new System.Windows.Media.FormattedText(
                    text,
                    System.Globalization.CultureInfo.CurrentCulture,
                    System.Windows.FlowDirection.LeftToRight,
                    new Typeface("Microsoft YaHei"),
                    12,
                    Brushes.DarkGray,
                    1.0);

                dc.DrawText(formattedText, new System.Windows.Point(
                    centerPoint.X - formattedText.Width / 2,
                    centerPoint.Y - formattedText.Height / 2));
            }

            // 渲染为位图
            var renderTarget = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            renderTarget.Render(visual);
            renderTarget.Freeze();

            return ConvertToBitmapImage(renderTarget);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 批量生成房间略缩图
    /// </summary>
    public Dictionary<long, BitmapImage?> GenerateThumbnails(IEnumerable<RoomData> rooms, int width = 300, int height = 200)
    {
        var results = new Dictionary<long, BitmapImage?>();

        foreach (var room in rooms)
        {
            results[room.ElementId] = GenerateRoomThumbnail(room.ElementId, width, height);
        }

        return results;
    }

    /// <summary>
    /// 导出略缩图到文件
    /// </summary>
    public bool ExportThumbnailToFile(long roomId, string filePath, int width = 300, int height = 200)
    {
        var thumbnail = GenerateRoomThumbnail(roomId, width, height);
        if (thumbnail == null) return false;

        try
        {
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(thumbnail));

            using var stream = new FileStream(filePath, FileMode.Create);
            encoder.Save(stream);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 计算边界框
    /// </summary>
    private BoundingBoxXYZ? CalculateBoundingBox(IList<IList<BoundarySegment>> boundarySegments)
    {
        if (boundarySegments == null || boundarySegments.Count == 0)
            return null;

        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;

        foreach (var loop in boundarySegments)
        {
            foreach (var segment in loop)
            {
                var curve = segment.GetCurve();
                var start = curve.GetEndPoint(0);
                var end = curve.GetEndPoint(1);

                minX = Math.Min(minX, Math.Min(start.X, end.X));
                minY = Math.Min(minY, Math.Min(start.Y, end.Y));
                maxX = Math.Max(maxX, Math.Max(start.X, end.X));
                maxY = Math.Max(maxY, Math.Max(start.Y, end.Y));
            }
        }

        var box = new BoundingBoxXYZ();
        box.Min = new XYZ(minX, minY, 0);
        box.Max = new XYZ(maxX, maxY, 0);
        return box;
    }

    /// <summary>
    /// 将 Revit 坐标转换为屏幕坐标
    /// </summary>
    private List<System.Windows.Point> ConvertToScreenPoints(
        IList<BoundarySegment> loop,
        BoundingBoxXYZ boundingBox,
        int width,
        int height)
    {
        var points = new List<System.Windows.Point>();

        double minX = boundingBox.Min.X;
        double minY = boundingBox.Min.Y;
        double rangeX = boundingBox.Max.X - minX;
        double rangeY = boundingBox.Max.Y - minY;

        // 保持宽高比
        double scale = Math.Min((width - 20) / rangeX, (height - 20) / rangeY);
        double offsetX = (width - rangeX * scale) / 2;
        double offsetY = (height - rangeY * scale) / 2;

        foreach (var segment in loop)
        {
            var curve = segment.GetCurve();
            var start = curve.GetEndPoint(0);

            double screenX = (start.X - minX) * scale + offsetX;
            double screenY = height - (start.Y - minY) * scale - offsetY; // Y 轴翻转

            points.Add(new System.Windows.Point(screenX, screenY));
        }

        return points;
    }

    /// <summary>
    /// 计算中心点（用于文字定位）
    /// </summary>
    private System.Windows.Point CalculateCenter(
        IList<IList<BoundarySegment>> boundarySegments,
        BoundingBoxXYZ boundingBox,
        int width,
        int height)
    {
        double minX = boundingBox.Min.X;
        double minY = boundingBox.Min.Y;
        double rangeX = boundingBox.Max.X - minX;
        double rangeY = boundingBox.Max.Y - minY;

        double scale = Math.Min((width - 20) / rangeX, (height - 20) / rangeY);
        double offsetX = (width - rangeX * scale) / 2;
        double offsetY = (height - rangeY * scale) / 2;

        // 计算所有点的平均位置
        double sumX = 0, sumY = 0;
        int count = 0;

        foreach (var loop in boundarySegments)
        {
            foreach (var segment in loop)
            {
                var curve = segment.GetCurve();
                var start = curve.GetEndPoint(0);
                sumX += start.X;
                sumY += start.Y;
                count++;
            }
        }

        if (count == 0)
            return new System.Windows.Point(width / 2, height / 2);

        double avgX = sumX / count;
        double avgY = sumY / count;

        double screenX = (avgX - minX) * scale + offsetX;
        double screenY = height - (avgY - minY) * scale - offsetY;

        return new System.Windows.Point(screenX, screenY);
    }

    /// <summary>
    /// 将 RenderTargetBitmap 转换为 BitmapImage
    /// </summary>
    private BitmapImage ConvertToBitmapImage(RenderTargetBitmap renderTarget)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(renderTarget));

        using var stream = new MemoryStream();
        encoder.Save(stream);
        stream.Position = 0;

        var bitmap = new BitmapImage();
        bitmap.BeginInit();
        bitmap.CacheOption = BitmapCacheOption.OnLoad;
        bitmap.StreamSource = stream;
        bitmap.EndInit();
        bitmap.Freeze();

        return bitmap;
    }
}
