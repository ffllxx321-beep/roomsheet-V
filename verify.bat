@echo off
chcp 65001 >nul
echo ========================================
echo 检查 Revit Addins 目录
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026

echo [1] 检查 RoomManager.dll
if exist "%REVIT_ADDINS%\RoomManager.dll" (
    echo     ✓ RoomManager.dll 存在
    for %%A in ("%REVIT_ADDINS%\RoomManager.dll") do echo     大小: %%~zA 字节
) else (
    echo     ✗ RoomManager.dll 不存在！
)

echo.
echo [2] 检查 RoomManager.addin
if exist "%REVIT_ADDINS%\RoomManager.addin" (
    echo     ✓ RoomManager.addin 存在
) else (
    echo     ✗ RoomManager.addin 不存在！
)

echo.
echo [3] 检查依赖 DLL...
set MISSING=0

if exist "%REVIT_ADDINS%\ACadSharp.dll" (
    echo     ✓ ACadSharp.dll
) else (
    echo     ✗ ACadSharp.dll 缺失！
    set MISSING=1
)

if exist "%REVIT_ADDINS%\EPPlus.dll" (
    echo     ✓ EPPlus.dll
) else (
    echo     ✗ EPPlus.dll 缺失！
    set MISSING=1
)

if exist "%REVIT_ADDINS%\itext.kernel.dll" (
    echo     ✓ itext.kernel.dll
) else (
    echo     ✗ itext.kernel.dll 缺失！
    set MISSING=1
)

if exist "%REVIT_ADDINS%\System.Drawing.Common.dll" (
    echo     ✓ System.Drawing.Common.dll
) else (
    echo     ✗ System.Drawing.Common.dll 缺失！
    set MISSING=1
)

echo.
echo [4] 列出所有 DLL 文件 (%REVIT_ADDINS%\*.dll):
dir /B "%REVIT_ADDINS%\*.dll" 2>nul

echo.
echo [5] 列出所有 .addin 文件:
dir /B "%REVIT_ADDINS%\*.addin" 2>nul

echo.
echo ========================================
if "%MISSING%"=="1" (
    echo ⚠️ 有依赖 DLL 缺失！
    echo.
    echo 请把 bin 目录下所有 .dll 文件都复制过去！
) else (
    echo ✓ 所有依赖 DLL 都存在
)
echo ========================================
pause
