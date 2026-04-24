using System.Windows;
using System.Windows.Controls;
using RoomManager.Models;
using RoomManager.Services;
using System.Linq;

namespace RoomManager.Views;

/// <summary>
/// 手动匹配对话框
/// </summary>
public partial class ManualMatchDialog : Window
{
    private readonly List<TextInRevit> _dwgTexts;
    private readonly MatchPreviewItem _item;

    public string SelectedText { get; private set; } = "";

    public ManualMatchDialog(List<TextInRevit> dwgTexts, MatchPreviewItem item)
    {
        InitializeComponent();
        _dwgTexts = dwgTexts;
        _item = item;

        DataContext = this;
        LoadTexts();
    }

    private void LoadTexts()
    {
        RoomInfoText.Text = $"{_item.RoomNumber} - {_item.RoomName}";
        CurrentMatchText.Text = _item.MatchedText;

        // 加载所有 DWG 文字
        foreach (var text in _dwgTexts)
        {
            TextListBox.Items.Add(new TextListItem
            {
                Content = text.Content,
                Position = $"({text.Position.X:F1}, {text.Position.Y:F1})",
                Layer = text.LayerName
            });
        }

        // 选中当前匹配
        if (!string.IsNullOrEmpty(_item.MatchedText))
        {
            var index = _dwgTexts.FindIndex(t => t.Content == _item.MatchedText);
            if (index >= 0)
            {
                TextListBox.SelectedIndex = index;
            }
        }
    }

    private void OnTextSelected(object sender, SelectionChangedEventArgs e)
    {
        if (TextListBox.SelectedItem is TextListItem item)
        {
            SelectedText = item.Content;
            PreviewText.Text = item.Content;
        }
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        var searchText = SearchTextBox.Text.ToLower();
        
        TextListBox.Items.Clear();
        
        var filtered = string.IsNullOrEmpty(searchText)
            ? _dwgTexts
            : _dwgTexts.Where(t => t.Content.ToLower().Contains(searchText));

        foreach (var text in filtered)
        {
            TextListBox.Items.Add(new TextListItem
            {
                Content = text.Content,
                Position = $"({text.Position.X:F1}, {text.Position.Y:F1})",
                Layer = text.LayerName
            });
        }
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(SelectedText))
        {
            MessageBox.Show("请选择一个文字。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
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

/// <summary>
/// 文字列表项
/// </summary>
public class TextListItem
{
    public string Content { get; set; } = "";
    public string Position { get; set; } = "";
    public string Layer { get; set; } = "";
}
