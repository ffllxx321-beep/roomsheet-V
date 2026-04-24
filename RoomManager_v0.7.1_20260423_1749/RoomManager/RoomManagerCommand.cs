using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RoomManager.Views;
using System.IO;
using System.Windows.Interop;

namespace RoomManager;

/// <summary>
/// Revit 外部命令 - 打开房间管理器
/// </summary>
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class RoomManagerCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uiApp = commandData.Application;
            var uiDoc = uiApp.ActiveUIDocument;
            var doc = uiDoc.Document;

            // 检查是否在平面视图
            var view = doc.ActiveView;
            var viewFamily = GetViewFamily(view);
            if (viewFamily != ViewFamily.FloorPlan && viewFamily != ViewFamily.CeilingPlan)
            {
                TaskDialog.Show("提示", "请在平面视图中使用此功能。");
                return Result.Cancelled;
            }

            // 循环：显示窗口 → 定位 → 重新显示
            while (true)
            {
                var window = new MainWindow();
                var helper = new WindowInteropHelper(window);
                helper.Owner = uiApp.MainWindowHandle;
                window.SetDocument(doc, view);
                window.ShowDialog();

                // 检查是否需要定位
                if (window.LocateElementId.HasValue)
                {
                    var elementId = new ElementId(window.LocateElementId.Value);
                    uiDoc.ShowElements(elementId);
                    // 继续循环，重新打开窗口
                }
                else
                {
                    break; // 正常关闭，退出循环
                }
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            // 记录完整异常信息
            var fullError = $"类型: {ex.GetType().FullName}\n消息: {ex.Message}\n堆栈: {ex.StackTrace}";
            if (ex.InnerException != null)
            {
                fullError += $"\n\n内部异常:\n类型: {ex.InnerException.GetType().FullName}\n消息: {ex.InnerException.Message}\n堆栈: {ex.InnerException.StackTrace}";
            }
            System.Diagnostics.Debug.WriteLine($"RoomManager 崩溃:\n{fullError}");
            try
            {
                var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RoomManager");
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                System.IO.File.WriteAllText(Path.Combine(logDir, "roommanager_crash.log"), fullError);
            }
            catch { }
            message = ex.Message;
            TaskDialog.Show("错误", fullError);
            return Result.Failed;
        }
    }

    /// <summary>
    /// 获取视图的 ViewFamily（兼容 Revit 2026）
    /// </summary>
    private ViewFamily GetViewFamily(View view)
    {
        try
        {
            // 尝试通过 ViewType 判断
            if (view.ViewType == ViewType.FloorPlan)
                return ViewFamily.FloorPlan;
            if (view.ViewType == ViewType.CeilingPlan)
                return ViewFamily.CeilingPlan;

            // 尝试通过 ViewFamily 属性判断（Revit 2024+）
            var viewType = view.GetType();
            var viewFamilyProp = viewType.GetProperty("ViewFamily");
            if (viewFamilyProp != null)
            {
                return (ViewFamily)viewFamilyProp.GetValue(view);
            }

            // 默认返回 Invalid
            return ViewFamily.Invalid;
        }
        catch
        {
            return ViewFamily.Invalid;
        }
    }
}

/// <summary>
/// Revit 外部命令 - 从 DWG 识别房间
/// </summary>
[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class DwgRecognitionCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var uiApp = commandData.Application;
            var uiDoc = uiApp.ActiveUIDocument;
            var doc = uiDoc.Document;

            // 检查是否在平面视图
            var view = doc.ActiveView;
            var viewFamily = GetViewFamily(view);
            if (viewFamily != ViewFamily.FloorPlan)
            {
                TaskDialog.Show("提示", "请在平面视图中使用此功能。");
                return Result.Cancelled;
            }

            // DWG 识别向导（支持框选循环）
            Outline? selectionArea = null;
            while (true)
            {
                var wizard = new Views.DwgRecognitionWizard(doc, view);
                if (selectionArea != null)
                    wizard.SelectionArea = selectionArea;
                var helper = new WindowInteropHelper(wizard);
                helper.Owner = uiApp.MainWindowHandle;
                wizard.ShowDialog();

                if (wizard.NeedsAreaSelection)
                {
                    // 用户要求框选 → 在 Revit 中框选
                    try
                    {
                        var pickBox = uiDoc.Selection.PickBox(Autodesk.Revit.UI.Selection.PickBoxStyle.Crossing, "请框选识别区域");
                        selectionArea = new Outline(pickBox.Min, pickBox.Max);
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        break; // 用户取消
                    }
                }
                else
                {
                    break; // 正常关闭
                }
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }

    /// <summary>
    /// 获取视图的 ViewFamily（兼容 Revit 2026）
    /// </summary>
    private ViewFamily GetViewFamily(View view)
    {
        try
        {
            if (view.ViewType == ViewType.FloorPlan)
                return ViewFamily.FloorPlan;
            if (view.ViewType == ViewType.CeilingPlan)
                return ViewFamily.CeilingPlan;

            var viewType = view.GetType();
            var viewFamilyProp = viewType.GetProperty("ViewFamily");
            if (viewFamilyProp != null)
            {
                return (ViewFamily)viewFamilyProp.GetValue(view);
            }

            return ViewFamily.Invalid;
        }
        catch
        {
            return ViewFamily.Invalid;
        }
    }
}

/// <summary>
/// Revit 应用程序 - 注册 Ribbon 面板
/// </summary>
public class RoomManagerApplication : IExternalApplication
{
    public Result OnStartup(UIControlledApplication application)
    {
        try
        {
            // 创建 Ribbon 面板
            string tabName = "绒姆徐特";
            application.CreateRibbonTab(tabName);

            RibbonPanel panel = application.CreateRibbonPanel(tabName, "房间管理");

            // 添加按钮
            string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            
            PushButtonData buttonData1 = new PushButtonData(
                "RoomManager",
                "房间管理器",
                assemblyPath,
                typeof(RoomManagerCommand).FullName);

            PushButton button1 = panel.AddItem(buttonData1) as PushButton;
            if (button1 != null)
            {
                button1.ToolTip = "打开房间信息管理器";
            }

            PushButtonData buttonData2 = new PushButtonData(
                "DwgRecognition",
                "DWG识别",
                assemblyPath,
                typeof(DwgRecognitionCommand).FullName);

            PushButton button2 = panel.AddItem(buttonData2) as PushButton;
            if (button2 != null)
            {
                button2.ToolTip = "从链接DWG识别房间";
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("错误", $"OnStartup 失败: {ex.Message}");
            return Result.Failed;
        }
    }

    public Result OnShutdown(UIControlledApplication application)
    {
        return Result.Succeeded;
    }
}
