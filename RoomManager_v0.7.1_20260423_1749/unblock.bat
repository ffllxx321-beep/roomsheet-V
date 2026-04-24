@echo off
echo ========================================
echo 解除 Windows 文件阻止
echo ========================================
echo.

set REVIT_ADDINS=C:\ProgramData\Autodesk\Revit\Addins\2026

echo 正在解除所有 DLL 的阻止状态...
powershell -Command "Get-ChildItem '%REVIT_ADDINS%\*.dll' | Unblock-File"

echo.
echo 正在解除 .addin 文件的阻止状态...
powershell -Command "Get-ChildItem '%REVIT_ADDINS%\*.addin' | Unblock-File"

echo.
echo ========================================
echo 完成！请重启 Revit 2026
echo ========================================
pause
