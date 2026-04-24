@echo off
chcp 65001 >nul
echo ========================================
echo RoomManager 安装脚本
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026

echo [步骤 1] 创建目录...
if not exist "%REVIT_ADDINS%" (
    mkdir "%REVIT_ADDINS%"
    echo     已创建: %REVIT_ADDINS%
) else (
    echo     目录已存在: %REVIT_ADDINS%
)

echo.
echo [步骤 2] 请输入 bin 目录路径...
echo    (例如: D:\RoomManager\bin\ 或直接按回车使用当前目录的 bin)
set /p BIN_PATH="bin 目录路径: "

if "%BIN_PATH%"=="" set BIN_PATH=bin

echo.
echo [步骤 3] 检查源文件...
if not exist "%BIN_PATH%\RoomManager.dll" (
    echo     [错误] 找不到 RoomManager.dll
    echo     请确认路径: %BIN_PATH%
    pause
    exit /b 1
)
echo     找到 RoomManager.dll ✓

echo.
echo [步骤 4] 复制文件到 Revit Addins 目录...

echo     复制 RoomManager.dll...
copy /Y "%BIN_PATH%\RoomManager.dll" "%REVIT_ADDINS%\" >nul
if exist "%REVIT_ADDINS%\RoomManager.dll" (
    echo     ✓ 成功
) else (
    echo     ✗ 失败
)

echo     复制 RoomManager.addin...
copy /Y "RoomManager\RoomManager.addin" "%REVIT_ADDINS%\" >nul
if exist "%REVIT_ADDINS%\RoomManager.addin" (
    echo     ✓ 成功
) else (
    echo     ✗ 失败
)

echo     复制依赖 DLL...
for %%f in ("%BIN_PATH%\*.dll") do (
    copy /Y "%%f" "%REVIT_ADDINS%\" >nul
)

echo.
echo [步骤 5] 验证安装...
echo.
echo     目标目录文件列表:
dir /B "%REVIT_ADDINS%\*.dll" 2>nul | find /c ".dll"
echo     个 DLL 文件

echo.
if exist "%REVIT_ADDINS%\RoomManager.dll" (
    echo     ✓ RoomManager.dll 已安装
    for %%A in ("%REVIT_ADDINS%\RoomManager.dll") do echo     大小: %%~zA 字节
) else (
    echo     ✗ RoomManager.dll 未找到！
)

if exist "%REVIT_ADDINS%\RoomManager.addin" (
    echo     ✓ RoomManager.addin 已安装
) else (
    echo     ✗ RoomManager.addin 未找到！
)

echo.
echo ========================================
echo 安装完成！请重启 Revit 2026
echo ========================================
pause