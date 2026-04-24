# Revit 插件"无法初始化附加模块"完整排查清单

## 错误信息
```
无法初始化附加模块"XXX"，因为程序集"XXX.dll"不存在。
```

---

## 🔍 排查清单（按可能性排序）

### ✅ 1. 文件位置检查

**检查项**：
- [ ] DLL 文件在 `C:\ProgramData\Autodesk\Revit\Addins\2026\` 目录
- [ ] .addin 文件在同一目录
- [ ] 文件名大小写正确（RoomManager.dll 不是 roommanager.dll）

**验证命令**：
```cmd
dir "C:\ProgramData\Autodesk\Revit\Addins\2026\RoomManager.*"
```

---

### ✅ 2. .addin 文件检查

**检查项**：
- [ ] `<Assembly>` 路径正确
- [ ] `<FullClassName>` 包含完整命名空间
- [ ] 文件编码是 UTF-8（不是 ANSI，不是 UTF-8 with BOM）
- [ ] XML 格式正确（没有语法错误）

**正确的 .addin 示例**：
```xml
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>插件名称</Name>
    <Assembly>插件名.dll</Assembly>  <!-- 相对路径或绝对路径 -->
    <FullClassName>命名空间.类名</FullClassName>
    <ClientId>GUID</ClientId>
    <VendorId>VENDOR</VendorId>
    <VendorDescription>Description</VendorDescription>
  </AddIn>
</RevitAddIns>
```

---

### ✅ 3. .NET 版本检查

**Revit 版本要求**：
| Revit 版本 | .NET 版本 |
|-----------|----------|
| Revit 2024 | .NET 4.8 |
| Revit 2025 | .NET 8 |
| Revit 2026 | .NET 8 |

**检查方法**：
- 打开 `.csproj` 文件
- 确认 `<TargetFramework>net8.0-windows</TargetFramework>`

---

### ✅ 4. Windows 安全阻止

**症状**：从网上下载的 DLL 被 Windows 阻止

**检查方法**：
- 右键 DLL → 属性 → 查看底部是否有"解除锁定"

**解决方法**：
```powershell
Get-ChildItem "C:\ProgramData\Autodesk\Revit\Addins\2026\*.dll" | Unblock-File
```

---

### ✅ 5. 依赖 DLL 缺失

**检查项**：
- [ ] 所有 NuGet 包的 DLL 都在 Addins 目录
- [ ] 没有缺失的依赖

**常见依赖**：
- ACadSharp.dll
- EPPlus.dll
- itext.kernel.dll
- System.Drawing.Common.dll
- BouncyCastle.Cryptography.dll

---

### ✅ 6. 编译配置检查

**检查项**：
- [ ] 使用 Release 配置编译
- [ ] 平台目标是 x64
- [ ] 输出目录正确

**正确的 .csproj 配置**：
```xml
<PropertyGroup>
  <TargetFramework>net8.0-windows</TargetFramework>
  <PlatformTarget>x64</PlatformTarget>
  <OutputPath>bin\</OutputPath>
  <CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>
</PropertyGroup>
```

---

### ✅ 7. 代码问题检查

**常见问题**：
- [ ] OnStartup 方法抛出异常
- [ ] 静态构造函数失败
- [ ] 类型初始化失败

**调试方法**：
在 OnStartup 中添加 try-catch：
```csharp
public Result OnStartup(UIControlledApplication application)
{
    try
    {
        // 你的代码
        return Result.Succeeded;
    }
    catch (Exception ex)
    {
        TaskDialog.Show("错误", ex.Message);
        return Result.Failed;
    }
}
```

---

### ✅ 8. Revit 日志检查

**日志位置**：
```
%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2026\Journals\
```

**检查方法**：
- 打开最新的 `.txt` 文件
- 搜索 "Error" 或 "Exception"
- 查找插件相关的错误信息

---

### ✅ 9. 权限问题

**检查项**：
- [ ] 有读取 Addins 目录的权限
- [ ] 有执行 DLL 的权限

**解决方法**：
以管理员身份运行 Revit

---

### ✅ 10. Revit 进程残留

**症状**：关闭 Revit 后后台还有进程

**检查方法**：
- 任务管理器 → 详细信息 → 查找 Revit.exe
- 结束所有 Revit 进程

---

## 🧪 最小化测试

如果以上都排查过了，创建一个最小化插件测试：

1. 创建新项目，只引用 RevitAPI 和 RevitAPIUI
2. 不添加任何第三方 NuGet 包
3. 只实现最简单的 OnStartup 和 IExternalCommand
4. 编译、部署、测试

如果最小化插件能加载，说明问题在代码或依赖上。

---

## 📞 仍然无法解决？

请提供以下信息：
1. Revit 日志中的完整错误信息
2. `C:\ProgramData\Autodesk\Revit\Addins\2026\` 目录下的文件列表
3. bin 目录下的文件列表
4. .addin 文件的完整内容
5. 编译输出窗口的信息
