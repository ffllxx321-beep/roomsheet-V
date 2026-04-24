using System.Windows;

namespace RoomManager.Views;

/// <summary>
/// 添加字段对话框
/// </summary>
public partial class AddFieldDialog : Window
{
    public string FieldName { get; set; } = string.Empty;
    public Models.FieldType FieldType { get; set; } = Models.FieldType.Text;
    public string? Options { get; set; }

    public AddFieldDialog()
    {
        InitializeComponent();
        DataContext = this;
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(FieldName))
        {
            MessageBox.Show("请输入字段名称。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (FieldType == Models.FieldType.Dropdown && string.IsNullOrWhiteSpace(Options))
        {
            MessageBox.Show("下拉字段必须提供选项（用逗号分隔）。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
