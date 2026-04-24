using System.Windows;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using RoomManager.Services;
using System.Collections.ObjectModel;
using WpfEllipse = System.Windows.Shapes.Ellipse;
using WpfRectangle = System.Windows.Shapes.Rectangle;

namespace RoomManager.Views;

/// <summary>
/// DWG 识别预览窗口
/// </summary>
public partial class DwgPreviewWindow : Window
{
    private readonly Document _document;
    private readonly List<TextInRevit> _dwgTexts;
    private readonly List<Room> _rooms;
    private List<RoomTextMatch> _matches = new();

    public ObservableCollection<MatchPreviewItem> PreviewItems { get; } = new();
    public bool IsConfirmed { get; private set; }

    public DwgPreviewWindow(Document document, List<TextInRevit> dwgTexts, List<Room> rooms)
    {
        InitializeComponent();
        _document = document;
        _dwgTexts = dwgTexts;
        _rooms = rooms;

        DataContext = this;
        LoadPreview();
    }

    private void LoadPreview()
    {
        // 执行匹配
        var service = new DwgRecognitionService();
        _matches = service.MatchTextsToRooms(_dwgTexts, _rooms);

        // 转换为预览项
        foreach (var match in _matches)
        {
            PreviewItems.Add(new MatchPreviewItem
            {
                RoomId = match.RoomId,
                RoomName = match.RoomName,
                RoomNumber = match.RoomNumber,
                OriginalText = match.MatchedText ?? "",
                MatchedText = match.MatchedText ?? "",
                Confidence = match.Confidence,
                Status = match.Status switch
                {
                    MatchStatus.Matched => "已匹配",
                    MatchStatus.NoText => "[!] 无文字",
                    MatchStatus.MultipleTexts => "[?] 多候选",
                    MatchStatus.LowConfidence => "[X] 低置信度",
                    _ => "[?] 未知"
                },
                IsManualEdit = false
            });
        }

        // 绘制预览图
        DrawPreview();
    }

    private void DrawPreview()
    {
        // 清空现有图形
        PreviewCanvas.Children.Clear();

        // 计算边界
        if (_dwgTexts.Count == 0) return;

        var minX = _dwgTexts.Min(t => t.Position.X);
        var maxX = _dwgTexts.Max(t => t.Position.X);
        var minY = _dwgTexts.Min(t => t.Position.Y);
        var maxY = _dwgTexts.Max(t => t.Position.Y);

        var width = maxX - minX;
        var height = maxY - minY;

        // 缩放因子
        var canvasWidth = PreviewCanvas.ActualWidth > 0 ? PreviewCanvas.ActualWidth : 600;
        var canvasHeight = PreviewCanvas.ActualHeight > 0 ? PreviewCanvas.ActualHeight : 400;
        var scale = Math.Min(canvasWidth / width, canvasHeight / height) * 0.9;

        // 绘制文字点
        foreach (var text in _dwgTexts)
        {
            var x = (text.Position.X - minX) * scale + 20;
            var y = canvasHeight - (text.Position.Y - minY) * scale - 20; // 翻转 Y 轴

            var ellipse = new WpfEllipse
            {
                Width = 6,
                Height = 6,
                Fill = Brushes.Blue,
                Opacity = 0.6
            };

            Canvas.SetLeft(ellipse, x - 3);
            Canvas.SetTop(ellipse, y - 3);
            PreviewCanvas.Children.Add(ellipse);
        }

        // 绘制房间中心点
        foreach (var room in _rooms)
        {
            // 简化：使用房间编号位置作为中心
            var location = room.Location as LocationPoint;
            if (location == null) continue;

            var x = (location.Point.X - minX) * scale + 20;
            var y = canvasHeight - (location.Point.Y - minY) * scale - 20;

            var rect = new WpfRectangle
            {
                Width = 10,
                Height = 10,
                Fill = Brushes.Green,
                Opacity = 0.6
            };

            Canvas.SetLeft(rect, x - 5);
            Canvas.SetTop(rect, y - 5);
            PreviewCanvas.Children.Add(rect);
        }
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        // 检查是否有低置信度匹配
        var lowConfidence = PreviewItems.Where(p => p.Confidence < 0.7).ToList();
        if (lowConfidence.Count > 0)
        {
            var result = MessageBox.Show(
                $"有 {lowConfidence.Count} 个低置信度匹配，是否继续？\n建议先手动调整这些匹配。",
                "警告",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;
        }

        IsConfirmed = true;
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        IsConfirmed = false;
        Close();
    }

    private void OnAutoMatch(object sender, RoutedEventArgs e)
    {
        // 重新自动匹配
        LoadPreview();
    }

    private void OnManualEdit(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is MatchPreviewItem item)
        {
            // 打开手动编辑对话框
            var dialog = new ManualMatchDialog(_dwgTexts, item) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                // 更新匹配
                item.MatchedText = dialog.SelectedText;
                item.Confidence = 1.0;
                item.Status = "手动匹配";
                item.IsManualEdit = true;
            }
        }
    }

    private void OnClearMatch(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is MatchPreviewItem item)
        {
            item.MatchedText = "";
            item.Confidence = 0;
            item.Status = "[!] 已清除";
            item.IsManualEdit = true;
        }
    }

    /// <summary>
    /// 获取最终匹配结果
    /// </summary>
    public List<TextRoomMatch> GetFinalMatches()
    {
        var results = new List<TextRoomMatch>();

        foreach (var item in PreviewItems)
        {
            if (!string.IsNullOrEmpty(item.MatchedText))
            {
                results.Add(new TextRoomMatch
                {
                    RoomId = item.RoomId,
                    RoomName = item.RoomName,
                    RoomNumber = item.RoomNumber,
                    MatchedText = item.MatchedText,
                    Confidence = item.Confidence,
                    Status = item.Confidence >= 0.7 ? MatchStatus.Matched : MatchStatus.LowConfidence
                });
            }
        }

        return results;
    }
}

/// <summary>
/// 匹配预览项
/// </summary>
public class MatchPreviewItem
{
    public long RoomId { get; set; }
    public string RoomName { get; set; } = "";
    public string RoomNumber { get; set; } = "";
    public string OriginalText { get; set; } = "";
    public string MatchedText { get; set; } = "";
    public double Confidence { get; set; }
    public string Status { get; set; } = "";
    public bool IsManualEdit { get; set; }
}

/// <summary>
/// 文字-房间匹配结果（已废弃，使用 DwgRecognitionService.RoomTextMatch）
/// </summary>
[Obsolete("使用 DwgRecognitionService.RoomTextMatch 替代")]
public class TextRoomMatch
{
    public long RoomId { get; set; }
    public string RoomName { get; set; } = "";
    public string RoomNumber { get; set; } = "";
    public string MatchedText { get; set; } = "";
    public double Confidence { get; set; }
    public MatchStatus Status { get; set; }
}
