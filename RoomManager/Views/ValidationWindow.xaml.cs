using System.Windows;
using RoomManager.Services;

namespace RoomManager.Views;

/// <summary>
/// 数据校验窗口交互逻辑
/// </summary>
public partial class ValidationWindow : Window
{
    private readonly ValidationReport _report;

    public ValidationWindow(ValidationReport report)
    {
        InitializeComponent();
        _report = report;
        DataContext = report;
    }

    private void OnExportReport(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "文本文件|*.txt|CSV 文件|*.csv",
            Title = "导出校验报告",
            FileName = $"校验报告_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                using var writer = new System.IO.StreamWriter(dialog.FileName);
                
                writer.WriteLine("=== 房间数据校验报告 ===");
                writer.WriteLine($"生成时间: {_report.ValidationTime:yyyy-MM-dd HH:mm:ss}");
                writer.WriteLine();
                writer.WriteLine($"总房间数: {_report.TotalRooms}");
                writer.WriteLine($"校验通过: {_report.ValidRoomCount}");
                writer.WriteLine($"错误数: {_report.ErrorCount}");
                writer.WriteLine($"警告数: {_report.WarningCount}");
                writer.WriteLine($"通过率: {_report.ValidRate:F1}%");
                writer.WriteLine();
                writer.WriteLine("=== 问题详情 ===");

                foreach (var (roomId, results) in _report.RoomResults)
                {
                    foreach (var result in results)
                    {
                        writer.WriteLine($"[{result.Severity}] 房间 {roomId} - {result.FieldName}: {result.Message}");
                    }
                }

                MessageBox.Show("报告已导出！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnClose(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
