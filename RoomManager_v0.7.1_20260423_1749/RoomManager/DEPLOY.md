# RoomManager 部署说明（完整版）

## ⚠️ 问题：找不到 RoomManager.dll

Revit 提示"程序集不存在"是因为 **DLL 没有复制到正确位置**。

---

## 📋 完整部署步骤

### 第一步：编译项目

1. 用 Visual Studio 2022 打开 `RoomManager.csproj`
2. 选择 **Release** 配置
3. 右键项目 → **生成**
4. 编译成功后，检查 `bin\` 目录下是否有 `RoomManager.dll`

### 第二步：复制文件到 Revit Addins 目录

**目标目录**：
```
C:\ProgramData\Autodesk\Revit\Addins\2026\
```

**必须复制的文件**（从 `bin\` 目录）：

```
✅ RoomManager.dll          ← 主程序
✅ RoomManager.addin        ← 插件注册文件（从项目根目录）
✅ ACadSharp.dll            ← DWG 读取
✅ EPPlus.dll               ← Excel 处理
✅ System.Drawing.Common.dll
✅ itext.kernel.dll         ← PDF 生成
✅ itext.layout.dll
✅ itext.io.dll
✅ itext.commons.dll
✅ BouncyCastle.Cryptography.dll  ← iText 依赖
```

### 第三步：验证部署

打开 `C:\ProgramData\Autodesk\Revit\Addins\2026\` 目录，应该看到：

```
RoomManager.dll
RoomManager.addin
ACadSharp.dll
EPPlus.dll
itext.kernel.dll
...（其他 DLL）
```

### 第四步：重启 Revit 2026

---

## 🚀 一键部署（推荐）

### 方法 1：使用部署脚本

1. 把 `deploy.bat` 放到项目根目录（和 `.csproj` 同级）
2. 编译项目
3. 双击运行 `deploy.bat`

### 方法 2：Visual Studio 自动部署

项目已配置 `DeployToRevit` Target，编译后自动复制。

**检查输出窗口**，应该看到：
```
已部署到: C:\ProgramData\Autodesk\Revit\Addins\2026\
```

---

## ❓ 常见问题

### Q: 编译后 bin 目录在哪？

A: 检查项目目录结构：
```
RoomManager/
├── bin/
│   ├── RoomManager.dll      ← 主 DLL
│   ├── ACadSharp.dll
│   ├── EPPlus.dll
│   └── ...（其他 DLL）
├── RoomManager.addin        ← 插件注册文件
└── RoomManager.csproj
```

### Q: 为什么 Revit 找不到 DLL？

A: 检查 `.addin` 文件中的路径是否正确：
```xml
<Assembly>C:\ProgramData\Autodesk\Revit\Addins\2026\RoomManager.dll</Assembly>
```

### Q: 如何确认 DLL 已复制？

A: 打开文件资源管理器，输入：
```
C:\ProgramData\Autodesk\Revit\Addins\2026
```
检查是否有 `RoomManager.dll`

---

## 📞 仍然有问题？

请提供以下信息：
1. `bin\` 目录下有哪些文件？
2. `C:\ProgramData\Autodesk\Revit\Addins\2026\` 目录下有哪些文件？
3. Revit 具体报错信息是什么？
