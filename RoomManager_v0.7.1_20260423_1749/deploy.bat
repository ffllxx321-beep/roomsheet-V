@echo off
echo ========================================
echo RoomManager Revit 2026 部署脚本
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026
set SOURCE_BIN=bin

echo 检查源目录...
if not exist "%SOURCE_BIN%\RoomManager.dll" (
    echo [错误] 请先编译项目！bin 目录下没有 RoomManager.dll
    pause
    exit /b 1
)

echo.
echo 创建目标目录...
if not exist "%REVIT_ADDINS%" mkdir "%REVIT_ADDINS%"

echo.
echo 复制 DLL 文件...
copy /Y "%SOURCE_BIN%\RoomManager.dll" "%REVIT_ADDINS%\"
copy /Y "%SOURCE_BIN%\ACadSharp.dll" "%REVIT_ADDINS%\"
copy /Y "%SOURCE_BIN%\EPPlus.dll" "%REVIT_ADDINS%\"
copy /Y "%SOURCE_BIN%\System.Drawing.Common.dll" "%REVIT_ADDINS%\"

echo 复制 iText DLL...
for %%f in (%SOURCE_BIN%\itext*.dll) do copy /Y "%%f" "%REVIT_ADDINS%\"

echo.
echo 复制 .addin 文件...
copy /Y "RoomManager.addin" "%REVIT_ADDINS%\"

echo.
echo ========================================
echo 部署完成！
echo 目标目录: %REVIT_ADDINS%
echo ========================================
echo.
echo 文件列表:
dir /B "%REVIT_ADDINS%\RoomManager*"
echo.
pause
