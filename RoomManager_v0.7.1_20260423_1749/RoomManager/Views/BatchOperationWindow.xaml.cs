using System.Windows;
using Autodesk.Revit.DB;
using RoomManager.Models;
using RoomManager.Services;
using System.Collections.Generic;

namespace RoomManager.Views;

/// <summary>
/// 批量操作窗口交互逻辑
/// </summary>
public partial class BatchOperationWindow : Window
{
    private readonly Document _document;
    private readonly List<RoomData> _selectedRooms;

    public BatchOperationWindow(Document document, List<RoomData> selectedRooms)
    {
        InitializeComponent();
        _document = document;
        _selectedRooms = selectedRooms;
        
        SelectedCountText.Text = $"已选中 {selectedRooms.Count} 间房间";
    }

    private void OnApplyRename(object sender, RoutedEventArgs e)
    {
        if (_selectedRooms.Count == 0)
        {
            MessageBox.Show("请先选择房间。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var result = MessageBox.Show(
            $"将对 {_selectedRooms.Count} 间房间应用重命名操作，是否继续？",
            "确认操作",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            var service = new BatchOperationService(_document);
            var roomIds = _selectedRooms.ConvertAll(r => r.ElementId);

            BatchResult? batchResult = null;

            // 正则替换
            var pattern = RegexPatternTextBox.Text;
            var replacement = RegexReplacementTextBox.Text;
            
            if (!string.IsNullOrEmpty(pattern))
            {
                batchResult = service.BatchRenameByRegex(roomIds, pattern, replacement);

                if (batchResult.HasError)
                {
                    MessageBox.Show($"操作失败: {batchResult.ErrorMessage}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            else
            {
                // 前缀/后缀
                var prefix = PrefixTextBox.Text;
                var suffix = SuffixTextBox.Text;

                if (!string.IsNullOrEmpty(prefix) || !string.IsNullOrEmpty(suffix))
                {
                    batchResult = service.BatchRenamePrefixSuffix(roomIds, prefix, suffix,
                        RemoveExistingCheckBox.IsChecked == true);

                    if (batchResult.HasError)
                    {
                        MessageBox.Show($"操作失败: {batchResult.ErrorMessage}", "错误",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }

            if (batchResult != null)
            {
                MessageBox.Show($"重命名完成！\n成功: {batchResult.SuccessCount}\n失败: {batchResult.FailCount}",
                    "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("请输入正则表达式或前缀/后缀。", "提示", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnApplyNumbering(object sender, RoutedEventArgs e)
    {
        if (_selectedRooms.Count == 0)
        {
            MessageBox.Show("请先选择房间。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var result = MessageBox.Show(
            $"将对 {_selectedRooms.Count} 间房间应用编号操作，是否继续？",
            "确认操作",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            var service = new BatchOperationService(_document);
            var roomIds = _selectedRooms.ConvertAll(r => r.ElementId);

            var template = NumberingTemplateTextBox.Text;
            var startIndex = int.TryParse(StartIndexTextBox.Text, out var start) ? start : 1;
            var step = int.TryParse(StepTextBox.Text, out var s) ? s : 1;

            var batchResult = service.BatchNumbering(roomIds, template, startIndex, step);

            if (batchResult.HasError)
            {
                MessageBox.Show($"操作失败: {batchResult.ErrorMessage}", "错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            MessageBox.Show($"编号完成！\n成功: {batchResult.SuccessCount}\n失败: {batchResult.FailCount}", 
                "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnApplyParameter(object sender, RoutedEventArgs e)
    {
        if (_selectedRooms.Count == 0)
        {
            MessageBox.Show("请先选择房间。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var paramName = ParameterNameTextBox.Text;
        if (string.IsNullOrWhiteSpace(paramName))
        {
            MessageBox.Show("请输入参数名称。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var result = MessageBox.Show(
            $"将对 {_selectedRooms.Count} 间房间设置参数 '{paramName}'，是否继续？",
            "确认操作",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            var service = new BatchOperationService(_document);
            var roomIds = _selectedRooms.ConvertAll(r => r.ElementId);
            var paramValue = ParameterValueTextBox.Text;

            var batchResult = service.BatchSetParameter(roomIds, paramName, paramValue);

            if (batchResult.HasError)
            {
                MessageBox.Show($"操作失败: {batchResult.ErrorMessage}", "错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            MessageBox.Show($"参数设置完成！\n成功: {batchResult.SuccessCount}\n失败: {batchResult.FailCount}", 
                "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnDeleteRooms(object sender, RoutedEventArgs e)
    {
        if (_selectedRooms.Count == 0)
        {
            MessageBox.Show("请先选择房间。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var result = MessageBox.Show(
            $"[警告] 即将删除 {_selectedRooms.Count} 间房间！\n此操作不可撤销，是否继续？",
            "危险操作确认",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes) return;

        // 二次确认
        result = MessageBox.Show(
            $"真的要删除 {_selectedRooms.Count} 间房间吗？\n输入 YES 确认删除",
            "最终确认",
            MessageBoxButton.YesNo,
            MessageBoxImage.Stop);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            var service = new BatchOperationService(_document);
            var roomIds = _selectedRooms.ConvertAll(r => r.ElementId);

            var batchResult = service.BatchDeleteRooms(roomIds);

            if (batchResult.HasError)
            {
                MessageBox.Show($"删除失败: {batchResult.ErrorMessage}", "错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            MessageBox.Show($"删除完成！\n成功: {batchResult.SuccessCount}\n失败: {batchResult.FailCount}", 
                "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"删除失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnClose(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
