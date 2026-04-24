@echo off
chcp 65001 >nul
echo ========================================
echo RoomManager 一键安装脚本
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026
set SCRIPT_DIR=%~dp0
set BIN_PATH=

echo [步骤 1] 创建目录...
if not exist "%REVIT_ADDINS%" (
    mkdir "%REVIT_ADDINS%"
    echo     已创建: %REVIT_ADDINS%
) else (
    echo     目录已存在: %REVIT_ADDINS%
)

echo.
echo [步骤 2] 自动查找可用编译产物...

if exist "%SCRIPT_DIR%bin\Release\net8.0-windows\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%bin\Release\net8.0-windows
)

if "%BIN_PATH%"=="" if exist "%SCRIPT_DIR%RoomManager\bin\Release\net8.0-windows\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%RoomManager\bin\Release\net8.0-windows
)

if "%BIN_PATH%"=="" if exist "%SCRIPT_DIR%bin\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%bin
)

if "%BIN_PATH%"=="" if exist "%SCRIPT_DIR%RoomManager\bin\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%RoomManager\bin
)

if "%BIN_PATH%"=="" (
    echo     未找到编译产物，尝试自动编译...
    where dotnet >nul 2>nul
    if errorlevel 1 (
        echo     [错误] 未检测到 dotnet 命令，无法自动编译。
        echo     请先安装 .NET 8 SDK，或把已编译好的 bin 目录放到安装脚本同级目录。
        pause
        exit /b 1
    )

    if exist "%SCRIPT_DIR%RoomManager\RoomManager.csproj" (
        echo     执行: dotnet build RoomManager\RoomManager.csproj -c Release
        dotnet build "%SCRIPT_DIR%RoomManager\RoomManager.csproj" -c Release
        if errorlevel 1 (
            echo     [错误] 自动编译失败，请查看上方日志。
            pause
            exit /b 1
        )
        set BIN_PATH=%SCRIPT_DIR%RoomManager\bin\Release\net8.0-windows
    ) else (
        echo     [错误] 未找到 RoomManager.csproj，无法自动编译。
        pause
        exit /b 1
    )
)

echo     使用目录: %BIN_PATH%

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

for %%f in ("%BIN_PATH%\*.dll") do (
    copy /Y "%%f" "%REVIT_ADDINS%\" >nul
)

if exist "%SCRIPT_DIR%RoomManager\RoomManager.addin" (
    copy /Y "%SCRIPT_DIR%RoomManager\RoomManager.addin" "%REVIT_ADDINS%\" >nul
)

if exist "%SCRIPT_DIR%RoomManager.addin" (
    copy /Y "%SCRIPT_DIR%RoomManager.addin" "%REVIT_ADDINS%\" >nul
)

if exist "%REVIT_ADDINS%\RoomManager.dll" (
    echo     ✓ DLL 复制成功
) else (
    echo     ✗ DLL 复制失败
)

if exist "%REVIT_ADDINS%\RoomManager.addin" (
    echo     ✓ Addin 复制成功
) else (
    echo     ✗ Addin 复制失败
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
echo 若插件未显示，请运行 verify.bat 进行诊断
echo ========================================
pause
