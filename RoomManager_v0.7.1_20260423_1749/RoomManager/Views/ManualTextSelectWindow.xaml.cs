using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RoomManager.Services;

namespace RoomManager.Views;

/// <summary>
/// 手动选择 DWG 文字窗口
/// </summary>
public partial class ManualTextSelectWindow : Window
{
    private readonly List<TextDisplayItem> _allItems = new();

    public string SelectedText { get; private set; } = "";

    public ManualTextSelectWindow(List<TextInRevit> texts, string roomNumber, string roomName)
    {
        InitializeComponent();
        RoomInfoRun.Text = $"{roomNumber} - {roomName}";

        // 构建显示列表（按内容去重，显示距离）
        var seen = new HashSet<string>();
        foreach (var text in texts.OrderBy(t => t.Content))
        {
            if (string.IsNullOrWhiteSpace(text.Content)) continue;
            var content = text.Content.Trim();
            if (seen.Contains(content)) continue;
            seen.Add(content);

            _allItems.Add(new TextDisplayItem
            {
                Content = content,
                LayerName = text.LayerName,
                Distance = ""
            });
        }

        TextListView.ItemsSource = _allItems;
    }

    private void OnSearchChanged(object sender, TextChangedEventArgs e)
    {
        var search = SearchBox.Text.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(search))
        {
            TextListView.ItemsSource = _allItems;
        }
        else
        {
            TextListView.ItemsSource = _allItems
                .Where(i => i.Content.ToLowerInvariant().Contains(search) ||
                            i.LayerName.ToLowerInvariant().Contains(search))
                .ToList();
        }
    }

    private void OnTextSelected(object sender, SelectionChangedEventArgs e)
    {
        if (TextListView.SelectedItem is TextDisplayItem item)
            SelectedText = item.Content;
    }

    private void OnTextDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (TextListView.SelectedItem is TextDisplayItem item)
        {
            SelectedText = item.Content;
            DialogResult = true;
            Close();
        }
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(SelectedText))
        {
            MessageBox.Show("请选择一个文字。", "提示");
            return;
        }
        DialogResult = true;
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}

public class TextDisplayItem
{
    public string Content { get; set; } = "";
    public string LayerName { get; set; } = "";
    public string Distance { get; set; } = "";
}
