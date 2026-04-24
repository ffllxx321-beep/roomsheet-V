using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using RoomManager.Models;
using RoomManager.Services;
using System.Linq;

namespace RoomManager.Views;

/// <summary>
/// 智能命名建议窗口交互逻辑
/// </summary>
public partial class SmartNamingWindow : Window
{
    private readonly Document _document;
    private readonly RoomData _roomData;
    private Room? _room;
    private List<NamingSuggestion> _suggestions = new();

    public SmartNamingWindow(Document document, RoomData roomData)
    {
        InitializeComponent();
        _document = document;
        _roomData = roomData;

        LoadRoomInfo();
        LoadSuggestions();
    }

    private void LoadRoomInfo()
    {
        CurrentRoomName.Text = _roomData.Name;
        CurrentRoomNumber.Text = _roomData.Number;
        CurrentRoomLevel.Text = _roomData.Level;
        CurrentRoomArea.Text = $"{_roomData.Area:F2} m²";

        // 获取 Revit Room 对象
        _room = _document.GetElement(new ElementId(_roomData.ElementId)) as Room;
    }

    private void LoadSuggestions()
    {
        if (_room == null)
        {
            MessageBox.Show("无法获取房间信息。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            var service = new SmartNamingService(_document);
            _suggestions = service.GetComprehensiveSuggestions(_room);

            SuggestionsList.ItemsSource = _suggestions;

            if (_suggestions.Count == 0)
            {
                MessageBox.Show("无法生成命名建议。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"生成建议失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnApplySuggestion(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is NamingSuggestion suggestion)
        {
            var result = MessageBox.Show(
                $"将房间名称修改为 \"{suggestion.SuggestedName}\"？\n\n原因: {suggestion.Reason}",
                "确认修改",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            try
            {
                var service = new SmartNamingService(_document);
                if (service.ApplySuggestion(_room!, suggestion))
                {
                    MessageBox.Show("名称已更新！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    // 更新显示
                    _roomData.Name = suggestion.SuggestedName;
                    CurrentRoomName.Text = suggestion.SuggestedName;
                }
                else
                {
                    MessageBox.Show("更新失败。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnRefresh(object sender, RoutedEventArgs e)
    {
        LoadSuggestions();
    }

    private void OnClose(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
