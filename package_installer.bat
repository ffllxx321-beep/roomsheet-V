@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========================================
echo RoomManager 安装向导打包脚本
echo ========================================
echo.

set SCRIPT_DIR=%~dp0
set OUT_DIR=%SCRIPT_DIR%release
set PAYLOAD_DIR=%OUT_DIR%\RoomManager_Payload
set ZIP_PATH=%OUT_DIR%\RoomManager_Payload.zip
set ISS_TEMPLATE=%SCRIPT_DIR%installer\RoomManagerSetup.iss
set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
set BIN_PATH=

if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"
if exist "%PAYLOAD_DIR%" rmdir /S /Q "%PAYLOAD_DIR%"
mkdir "%PAYLOAD_DIR%"

echo [1/5] 自动查找编译产物...
if exist "%SCRIPT_DIR%bin\Release\net8.0-windows\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%bin\Release\net8.0-windows
)
if "!BIN_PATH!"=="" if exist "%SCRIPT_DIR%RoomManager\bin\Release\net8.0-windows\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%RoomManager\bin\Release\net8.0-windows
)
if "!BIN_PATH!"=="" if exist "%SCRIPT_DIR%bin\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%bin
)
if "!BIN_PATH!"=="" if exist "%SCRIPT_DIR%RoomManager\bin\RoomManager.dll" (
    set BIN_PATH=%SCRIPT_DIR%RoomManager\bin
)

if "!BIN_PATH!"=="" (
    echo     未找到编译产物，尝试 dotnet build -c Release...
    where dotnet >nul 2>nul
    if errorlevel 1 (
        echo     [错误] 未检测到 dotnet，无法自动编译。
        echo     请先编译项目后再执行此脚本。
        exit /b 1
    )

    if not exist "%SCRIPT_DIR%RoomManager\RoomManager.csproj" (
        echo     [错误] 未找到 RoomManager.csproj。
        exit /b 1
    )

    dotnet build "%SCRIPT_DIR%RoomManager\RoomManager.csproj" -c Release
    if errorlevel 1 (
        echo     [错误] 自动编译失败，请检查编译日志。
        exit /b 1
    )
    set BIN_PATH=%SCRIPT_DIR%RoomManager\bin\Release\net8.0-windows
)

echo     使用目录: !BIN_PATH!

echo [2/5] 复制插件文件到 Payload...
copy /Y "!BIN_PATH!\*.dll" "%PAYLOAD_DIR%\" >nul

if exist "%SCRIPT_DIR%RoomManager\RoomManager.addin" (
    copy /Y "%SCRIPT_DIR%RoomManager\RoomManager.addin" "%PAYLOAD_DIR%\" >nul
) else if exist "%SCRIPT_DIR%RoomManager.addin" (
    copy /Y "%SCRIPT_DIR%RoomManager.addin" "%PAYLOAD_DIR%\" >nul
) else (
    echo     [错误] 未找到 RoomManager.addin。
    exit /b 1
)

copy /Y "%SCRIPT_DIR%verify.bat" "%PAYLOAD_DIR%\" >nul
copy /Y "%SCRIPT_DIR%install.bat" "%PAYLOAD_DIR%\" >nul
copy /Y "%SCRIPT_DIR%README.md" "%PAYLOAD_DIR%\" >nul

if not exist "%PAYLOAD_DIR%\RoomManager.dll" (
    echo     [错误] Payload 缺少 RoomManager.dll。
    exit /b 1
)

echo [3/5] 生成备用 ZIP 安装包...
if exist "%ZIP_PATH%" del /Q "%ZIP_PATH%"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Compress-Archive -Path '%PAYLOAD_DIR%\*' -DestinationPath '%ZIP_PATH%' -Force"
if exist "%ZIP_PATH%" (
    echo     ✓ 已生成: %ZIP_PATH%
) else (
    echo     ⚠️ 未生成 ZIP（PowerShell 不可用时可忽略）。
)

echo [4/5] 构建安装向导 EXE (Inno Setup)...
if not exist "%ISS_TEMPLATE%" (
    echo     [错误] 未找到安装脚本模板: %ISS_TEMPLATE%
    exit /b 1
)

if exist "%ISCC%" (
    "%ISCC%" "/DPayloadDir=%PAYLOAD_DIR%" "%ISS_TEMPLATE%"
    if errorlevel 1 (
        echo     [错误] Inno Setup 编译失败。
        exit /b 1
    )
    echo     ✓ 安装向导已生成（见 installer 输出目录）
) else (
    echo     ⚠️ 未检测到 Inno Setup 6:
    echo        %ISCC%
    echo     请安装 Inno Setup 6 后重新运行，即可生成 Setup.exe。
)

echo [5/5] 完成
echo.
echo Payload 目录: %PAYLOAD_DIR%
if exist "%ZIP_PATH%" echo ZIP 包: %ZIP_PATH%
echo.
pause
