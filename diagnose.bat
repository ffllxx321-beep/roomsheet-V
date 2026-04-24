@echo off
chcp 65001 >nul
echo ========================================
echo RoomManager 完整诊断
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026

echo [检查 1] 目录是否存在
if exist "%REVIT_ADDINS%" (
    echo     ✓ %REVIT_ADDINS%
) else (
    echo     ✗ 目录不存在！
    goto :end
)

echo.
echo [检查 2] RoomManager.dll
if exist "%REVIT_ADDINS%\RoomManager.dll" (
    echo     ✓ 文件存在
    for %%A in ("%REVIT_ADDINS%\RoomManager.dll") do echo     大小: %%~zA 字节
    for %%A in ("%REVIT_ADDINS%\RoomManager.dll") do echo     修改时间: %%~tA
) else (
    echo     ✗ 文件不存在！
)

echo.
echo [检查 3] RoomManager.addin
if exist "%REVIT_ADDINS%\RoomManager.addin" (
    echo     ✓ 文件存在
    echo     内容:
    type "%REVIT_ADDINS%\RoomManager.addin"
) else (
    echo     ✗ 文件不存在！
)

echo.
echo [检查 4] 所有 DLL 文件数量
for /f %%i in ('dir /B "%REVIT_ADDINS%\*.dll" 2^>nul ^| find /c /v ""') do set DLL_COUNT=%%i
echo     共 %DLL_COUNT% 个 DLL 文件

echo.
echo [检查 5] 关键依赖
for %%d in (ACadSharp EPPlus itext.kernel itext.layout System.Drawing.Common BouncyCastle.Cryptography) do (
    if exist "%REVIT_ADDINS%\%%d.dll" (
        echo     ✓ %%d.dll
    ) else (
        echo     ✗ %%d.dll 缺失
    )
)

echo.
echo [检查 6] 文件权限
icacls "%REVIT_ADDINS%\RoomManager.dll" 2>nul | find "RoomManager.dll"

echo.
echo [检查 7] .addin 文件编码
file "%REVIT_ADDINS%\RoomManager.addin" 2>nul || echo     (无法检测编码，请用记事本打开确认是 UTF-8)

echo.
echo ========================================
echo 诊断完成
echo ========================================
echo.
echo 如果所有文件都存在，请检查:
echo 1. .addin 文件是否是 UTF-8 编码（用记事本另存为 UTF-8）
echo 2. Revit 是否完全关闭后重新打开
echo 3. 是否有多个 Revit 进程在后台运行
echo.

:end
pause
