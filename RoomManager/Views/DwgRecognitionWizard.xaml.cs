using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using RoomManager.Services;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace RoomManager.Views;

/// <summary>
/// DWG 识别向导窗口
/// </summary>
public partial class DwgRecognitionWizard : Window
{
    private readonly Document _document;
    private readonly View _view;
    private readonly RevitRoomService _roomService;
    private readonly DwgRecognitionService _dwgService;
    private List<TextInRevit> _extractedTexts = new();
    private string[]? _selectedLayerFilter;

    public ObservableCollection<CADLinkInfo> CADLinks { get; } = new();
    public ObservableCollection<DwgMatchItem> MatchResults { get; } = new();
    public ObservableCollection<LayerItem> Layers { get; } = new();

    public CADLinkInfo? SelectedCADLink { get; set; }
    public bool AutoPlaceRooms { get; set; } = true;
    public bool NeedsAreaSelection { get; set; }
    public Outline? SelectionArea { get; set; }

    public DwgRecognitionWizard(Document document, View view)
    {
        InitializeComponent();
        DataContext = this;
        
        _document = document;
        _view = view;
        _roomService = new RevitRoomService(document);
        _dwgService = new DwgRecognitionService();

        LoadCADLinks();
    }

    private void LoadCADLinks()
    {
        var links = _roomService.GetLinkedCADLinks(_view);
        foreach (var link in links)
            CADLinks.Add(link);

        if (CADLinks.Count == 0)
        {
            MessageBox.Show("当前视图中没有链接的 CAD 图纸，请先链接 DWG 文件。", "提示", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            Close();
        }
    }

    /// <summary>
    /// 扫描 DWG 图层
    /// </summary>
    private void OnScanLayers(object sender, RoutedEventArgs e)
    {
        if (SelectedCADLink == null)
        {
            MessageBox.Show("请先选择一个 CAD 图纸。", "提示");
            return;
        }

        try
        {
            StatusText.Text = "正在扫描图层...";
            var texts = _dwgService.ExtractTextsFromDwg(SelectedCADLink.FilePath);

            Layers.Clear();
            var layerGroups = texts.GroupBy(t => t.LayerName).OrderBy(g => g.Key);
            foreach (var group in layerGroups)
            {
                Layers.Add(new LayerItem
                {
                    Name = group.Key,
                    TextCount = group.Count(),
                    IsSelected = false
                });
            }

            LayerListBox.ItemsSource = Layers;
            StatusText.Text = $"发现 {Layers.Count} 个图层，共 {texts.Count} 个文字。请选择房间名所在的图层。";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"扫描图层失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnSelectAllLayers(object sender, RoutedEventArgs e)
    {
        foreach (var layer in Layers) layer.IsSelected = true;
    }

    private void OnDeselectAllLayers(object sender, RoutedEventArgs e)
    {
        foreach (var layer in Layers) layer.IsSelected = false;
    }

    private void OnLayerCheckChanged(object sender, RoutedEventArgs e) { }

    /// <summary>
    /// 预览识别
    /// </summary>
    private void OnPreviewClick(object sender, RoutedEventArgs e)
    {
        if (SelectedCADLink == null)
        {
            MessageBox.Show("请选择要识别的 CAD 图纸。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            StatusText.Text = "正在提取文字...";

            // 获取选中的图层（空=全部）
            var selectedLayers = Layers.Where(l => l.IsSelected).Select(l => l.Name).ToArray();
            string[]? layerFilter = selectedLayers.Length > 0 ? selectedLayers : null;
            _selectedLayerFilter = layerFilter;

            // 提取 DWG 文字
            var dwgTexts = _dwgService.ExtractTextsFromDwg(SelectedCADLink.FilePath, layerFilter);
            _extractedTexts = _dwgService.ConvertToRevitCoordinates(dwgTexts, SelectedCADLink.Transform);

            StatusText.Text = $"提取到 {_extractedTexts.Count} 个文字，正在获取房间...";

            // 获取房间
            var rooms = GetRooms();

            // 如果没有房间且勾选了自动放置
            if (rooms.Count == 0 && AutoPlaceRooms)
            {
                StatusText.Text = "未找到房间，正在自动放置...";
                int placedCount = AutoPlaceRoomsInLevel();
                rooms = GetRooms(); // 重新获取
                StatusText.Text = $"自动放置了 {placedCount} 个房间";
            }

            if (rooms.Count == 0)
            {
                MessageBox.Show(
                    "当前楼层没有找到可放置房间的闭合空间。\n\n" +
                    "可能原因：\n" +
                    "1. DWG 中的墙线不是 Revit 的 Room-Bounding 元素\n" +
                    "2. 需要先在 Revit 中绘制墙体或房间分隔线\n\n" +
                    "建议：先手动放置房间，再用本工具匹配名称。", 
                    "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            StatusText.Text = $"找到 {rooms.Count} 个房间，正在匹配...";

            // 匹配
            var matches = _dwgService.MatchTextsToRooms(_extractedTexts, rooms);

            // 显示结果
            MatchResults.Clear();
            int matched = 0, noText = 0;
            foreach (var match in matches)
            {
                var item = new DwgMatchItem
                {
                    RoomId = match.RoomId,
                    RoomName = match.RoomName,
                    RoomNumber = match.RoomNumber,
                    MatchedText = match.MatchedText,
                    Confidence = match.Confidence,
                    Status = match.Status,
                    StatusText = match.Status switch
                    {
                        MatchStatus.Matched => "已匹配",
                        MatchStatus.NoText => "未找到",
                        MatchStatus.MultipleTexts => "多候选",
                        MatchStatus.LowConfidence => "低置信",
                        _ => "未知"
                    }
                };
                MatchResults.Add(item);
                if (match.Status == MatchStatus.Matched) matched++;
                if (match.Status == MatchStatus.NoText) noText++;
            }

            MatchSummary.Text = $"共 {matches.Count} 个房间: {matched} 已匹配, {noText} 未找到";
            StatusText.Text = "识别完成。可以手动修改识别名称后点击「确认并填充」。";
            ResultTab.IsSelected = true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"识别失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "识别失败";
        }
    }

    private List<Room> GetRooms()
    {
        var targetPhase = GetTargetPhase();
        var allRooms = new FilteredElementCollector(_document)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .Cast<Room>()
            .Where(r => r.Area > 0 && (targetPhase == null || r.CreatedPhaseId == targetPhase.Id))
            .ToList();

        if (_view.GenLevel != null)
        {
            var levelId = _view.GenLevel.Id;
            allRooms = allRooms.Where(r => r.Level?.Id == levelId).ToList();
        }

        return allRooms;
    }

    /// <summary>
    /// 自动放置房间：先尝试 NewRooms2（利用 Revit 墙体），再用 DWG 文字位置补充
    /// </summary>
    private int AutoPlaceRoomsInLevel()
    {
        int placed = 0;
        if (_view.GenLevel == null) return 0;
        if (SelectedCADLink == null) return 0;

        var level = _view.GenLevel;
        var phase = GetTargetPhase();
        if (phase == null) return 0;

        using var transaction = new Transaction(_document, "自动放置房间");
        transaction.Start();
        var options = transaction.GetFailureHandlingOptions();
        options.SetFailuresPreprocessor(new RoomWarningSwallower());
        transaction.SetFailureHandlingOptions(options);

        try
        {
            // 方式1: 用 Revit 墙体自动放置（如果有 Revit 墙围合的空间）
            try
            {
                var existingRooms = new FilteredElementCollector(_document)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>()
                    .Where(r => r.Level?.Id == level.Id && r.CreatedPhaseId == phase.Id)
                    .ToList();

                // 当前楼层+阶段已有房间时，不再重复调用 NewRooms2，避免“同一围合区多个房间”警告
                if (existingRooms.Count == 0)
                {
                    var newRoomIds = _document.Create.NewRooms2(level, phase);
                    foreach (ElementId roomId in newRoomIds)
                    {
                        if (_document.GetElement(roomId) is Room createdRoom)
                        {
                            if (IsPlacedRoom(createdRoom))
                            {
                                placed++;
                            }
                            else
                            {
                                // 清理无面积/未正确放置的房间，避免预览重复点击后堆积
                                _document.Delete(roomId);
                            }
                        }
                    }
                }
            }
            catch { }

            // 方式2: 用 DWG 闭合区域中心放置（优先满足“闭合区域无房间先创建”）
            if (!string.IsNullOrWhiteSpace(SelectedCADLink.FilePath))
            {
                try
                {
                    var dwgTexts = _dwgService.ExtractTextsFromDwg(SelectedCADLink.FilePath, _selectedLayerFilter);
                    var regions = _dwgService.ExtractClosedRegions(SelectedCADLink.FilePath, _selectedLayerFilter)
                        .Where(r => r.Area > 1.0)
                        .ToList();
                    var regionMatches = _dwgService.MatchTextsToRegions(dwgTexts, regions, 30.0);

                    foreach (var regionMatch in regionMatches)
                    {
                        if (string.IsNullOrWhiteSpace(regionMatch.MatchedText)) continue;

                        var dwgCenter = new XYZ(regionMatch.Region.CenterX, regionMatch.Region.CenterY, 0);
                        var revitCenter = SelectedCADLink.Transform.OfPoint(dwgCenter);
                        var point = new UV(revitCenter.X, revitCenter.Y);
                        var testPoint = new XYZ(revitCenter.X, revitCenter.Y, level.ProjectElevation + 1.0);

                        if (PointHasExistingRoom(testPoint, level.Id)) continue;

                        try
                        {
                            var room = _document.Create.NewRoom(level, point);
                            if (room != null)
                            {
                                if (IsPlacedRoom(room))
                                {
                                    room.Name = regionMatch.MatchedText.Trim();
                                    placed++;
                                }
                                else
                                {
                                    _document.Delete(room.Id);
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            // 方式3: 用 DWG 文字位置放置房间（在文字位置没有房间的地方）
            if (_extractedTexts.Count > 0)
            {
                double levelZ = level.ProjectElevation;

                foreach (var text in _extractedTexts)
                {
                    if (string.IsNullOrWhiteSpace(text.Content)) continue;

                    // 创建放置点（使用文字的 XY + 楼层的 Z）
                    var point = new UV(text.Position.X, text.Position.Y);

                    // 检查这个点是否已经在某个房间内
                    var testPoint = new XYZ(text.Position.X, text.Position.Y, levelZ + 1.0);
                    if (PointHasExistingRoom(testPoint, level.Id)) continue;

                    // 在文字位置放置新房间
                    try
                    {
                        var newRoom = _document.Create.NewRoom(level, point);
                        if (newRoom != null)
                        {
                            if (IsPlacedRoom(newRoom))
                            {
                                newRoom.Name = text.Content.Trim();
                                placed++;
                            }
                            else
                            {
                                _document.Delete(newRoom.Id);
                            }
                        }
                    }
                    catch { } // 放不了就跳过（可能不在闭合区域内）
                }
            }

            transaction.Commit();
        }
        catch
        {
            if (transaction.HasStarted())
                transaction.RollBack();
        }

        return placed;
    }

    private static bool IsPlacedRoom(Room room)
    {
        try
        {
            return room.Area > 0 && room.Location is LocationPoint;
        }
        catch
        {
            return false;
        }
    }

    private Phase? GetTargetPhase()
    {
        try
        {
            var viewPhaseParam = _view.get_Parameter(BuiltInParameter.VIEW_PHASE);
            if (viewPhaseParam != null)
            {
                var phaseId = viewPhaseParam.AsElementId();
                if (phaseId != null && phaseId != ElementId.InvalidElementId)
                {
                    if (_document.GetElement(phaseId) is Phase viewPhase)
                    {
                        return viewPhase;
                    }
                }
            }
        }
        catch { }

        // 回退：使用最后一个阶段
        try
        {
            var phases = _document.Phases;
            if (phases.Size > 0)
                return phases.get_Item(phases.Size - 1);
        }
        catch { }

        return null;
    }

    private bool PointHasExistingRoom(XYZ point, ElementId levelId)
    {
        try
        {
            var existingRooms = new FilteredElementCollector(_document)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r.Area > 0 && r.Level?.Id == levelId);

            foreach (var existingRoom in existingRooms)
            {
                try
                {
                    if (existingRoom.IsPointInRoom(point))
                    {
                        return true;
                    }
                }
                catch { }
            }
        }
        catch { }

        return false;
    }

    /// <summary>
    /// 手动选择文字
    /// </summary>
    private void OnManualSelect(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not DwgMatchItem item) return;

        // 弹出文字选择列表
        var selectWindow = new ManualTextSelectWindow(_extractedTexts, item.RoomNumber, item.RoomName);
        selectWindow.Owner = this;
        if (selectWindow.ShowDialog() == true)
        {
            item.MatchedText = selectWindow.SelectedText;
            item.Confidence = 1.0;
            item.StatusText = "手动选择";
            ResultGrid.Items.Refresh();
        }
    }

    private void OnClearMatch(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not DwgMatchItem item) return;
        item.MatchedText = "";
        item.Confidence = 0;
        item.StatusText = "已清除";
        ResultGrid.Items.Refresh();
    }

    /// <summary>
    /// 确认并填充
    /// </summary>
    private void OnApplyClick(object sender, RoutedEventArgs e)
    {
        if (MatchResults.Count == 0)
        {
            MessageBox.Show("请先预览识别结果。", "提示");
            return;
        }

        var toUpdate = MatchResults.Where(m => !string.IsNullOrEmpty(m.MatchedText)).ToList();
        if (toUpdate.Count == 0)
        {
            MessageBox.Show("没有可更新的匹配结果。", "提示");
            return;
        }

        var result = MessageBox.Show(
            $"即将更新 {toUpdate.Count} 个房间的名称，是否继续？",
            "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            using var transaction = new Transaction(_document, "从 DWG 更新房间名称");
            transaction.Start();

            int success = 0;
            foreach (var item in toUpdate)
            {
                var elementId = new ElementId(item.RoomId);
                if (_document.GetElement(elementId) is Room room)
                {
                    room.Name = item.MatchedText;
                    success++;
                }
            }

            transaction.Commit();
            MessageBox.Show($"成功更新 {success} 个房间的名称。", "完成");
            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"更新失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}

internal class RoomWarningSwallower : IFailuresPreprocessor
{
    public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
    {
        var messages = failuresAccessor.GetFailureMessages();
        foreach (var msg in messages)
        {
            if (msg.GetSeverity() == FailureSeverity.Warning)
            {
                failuresAccessor.DeleteWarning(msg);
            }
        }
        return FailureProcessingResult.Continue;
    }
}

/// <summary>
/// DWG 匹配结果项
/// </summary>
public class DwgMatchItem
{
    public long RoomId { get; set; }
    public string RoomName { get; set; } = "";
    public string RoomNumber { get; set; } = "";
    public string MatchedText { get; set; } = "";
    public double Confidence { get; set; }
    public MatchStatus Status { get; set; }
    public string StatusText { get; set; } = "";
}

/// <summary>
/// 图层项
/// </summary>
public class LayerItem : INotifyPropertyChanged
{
    public string Name { get; set; } = "";
    public int TextCount { get; set; }
    
    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); }
    }
    
    public event PropertyChangedEventHandler? PropertyChanged;
}

/// <summary>
/// 识别模式
/// </summary>
public enum RecognitionMode
{
    CurrentLevel,
    SelectionArea
}
