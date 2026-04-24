using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using RoomManager.Models;
using RoomManager.ViewModels;

namespace RoomManager.Views;

/// <summary>
/// 字段管理器窗口 — 管理参数的显示/隐藏
/// </summary>
public partial class FieldManagerWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly List<FieldVisibilityItem> _items = new();

    public FieldManagerWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BuildFieldList();
    }

    /// <summary>
    /// 从当前房间数据构建字段列表
    /// </summary>
    private void BuildFieldList()
    {
        _items.Clear();

        // 从所有房间的参数中收集不重复的字段名
        var allParamNames = new Dictionary<string, (string displayName, string groupName, bool isReadOnly)>();

        foreach (var room in _viewModel.Rooms)
        {
            foreach (var param in room.AllParameters)
            {
                if (!allParamNames.ContainsKey(param.Name))
                {
                    allParamNames[param.Name] = (param.DisplayName, param.GroupName, param.IsReadOnly);
                }
            }
        }

        // 构建可视项
        foreach (var (name, info) in allParamNames.OrderBy(x => x.Value.groupName).ThenBy(x => x.Value.displayName))
        {
            _items.Add(new FieldVisibilityItem
            {
                Name = name,
                DisplayName = info.displayName,
                GroupName = info.groupName,
                IsReadOnly = info.isReadOnly,
                IsVisible = _viewModel.FieldManager.IsFieldVisible(name)
            });
        }

        FieldsList.ItemsSource = _items;

        // 按组分组
        var view = CollectionViewSource.GetDefaultView(_items);
        view.GroupDescriptions.Add(new PropertyGroupDescription("GroupName"));

        UpdateCount();
    }

    private void UpdateCount()
    {
        var visible = _items.Count(i => i.IsVisible);
        CountLabel.Text = $"{visible}/{_items.Count} 个字段可见";
    }

    private void OnVisibilityChanged(object sender, RoutedEventArgs e)
    {
        UpdateCount();
    }

    private void OnShowAll(object sender, RoutedEventArgs e)
    {
        foreach (var item in _items) item.IsVisible = true;
        UpdateCount();
    }

    private void OnHideAll(object sender, RoutedEventArgs e)
    {
        foreach (var item in _items) item.IsVisible = false;
        UpdateCount();
    }

    private void OnShowEditableOnly(object sender, RoutedEventArgs e)
    {
        foreach (var item in _items) item.IsVisible = !item.IsReadOnly;
        UpdateCount();
    }

    private void OnReset(object sender, RoutedEventArgs e)
    {
        foreach (var item in _items) item.IsVisible = true;
        UpdateCount();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e)
    {
        // 保存到 FieldManager
        _viewModel.FieldManager.ResetAll();
        foreach (var item in _items.Where(i => !i.IsVisible))
        {
            _viewModel.FieldManager.SetFieldVisibility(item.Name, false);
        }
        _viewModel.FieldManager.SaveConfig();

        // 刷新详情面板
        var current = _viewModel.SelectedRoom;
        _viewModel.SelectedRoom = null;
        _viewModel.SelectedRoom = current;

        DialogResult = true;
        Close();
    }
}
