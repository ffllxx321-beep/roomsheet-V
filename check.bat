@echo off
echo ========================================
echo RoomManager 部署检查脚本
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026

echo [1] 检查目标目录是否存在...
if exist "%REVIT_ADDINS%" (
    echo     ✓ 目录存在
) else (
    echo     ✗ 目录不存在！
)

echo.
echo [2] 检查 RoomManager.dll...
if exist "%REVIT_ADDINS%\RoomManager.dll" (
    echo     ✓ RoomManager.dll 存在
    for %%A in ("%REVIT_ADDINS%\RoomManager.dll") do echo     大小: %%~zA 字节
) else (
    echo     ✗ RoomManager.dll 不存在！
)

echo.
echo [3] 检查 RoomManager.addin...
if exist "%REVIT_ADDINS%\RoomManager.addin" (
    echo     ✓ RoomManager.addin 存在
    echo     内容:
    type "%REVIT_ADDINS%\RoomManager.addin"
) else (
    echo     ✗ RoomManager.addin 不存在！
)

echo.
echo [4] 检查依赖 DLL...
for %%d in (ACadSharp EPPlus itext.kernel System.Drawing.Common) do (
    if exist "%REVIT_ADDINS%\%%d.dll" (
        echo     ✓ %%d.dll
    ) else (
        echo     ✗ %%d.dll 缺失！
    )
)

echo.
echo [5] 列出所有文件...
dir /B "%REVIT_ADDINS%\*.dll" 2>nul
dir /B "%REVIT_ADDINS%\*.addin" 2>nul

echo.
echo ========================================
pause
