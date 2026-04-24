# 检查 Revit 安装路径
Write-Host "=== 检查 Revit 安装 ===" -ForegroundColor Cyan

$revitVersions = @("2026", "2025", "2024", "2023", "2022")
$foundRevit = @()

foreach ($version in $revitVersions) {
    $revitPath = "C:\Program Files\Autodesk\Revit $version"
    if (Test-Path $revitPath) {
        $apiPath = Join-Path $revitPath "RevitAPI.dll"
        if (Test-Path $apiPath) {
            Write-Host "✅ 找到 Revit $version" -ForegroundColor Green
            Write-Host "   路径: $revitPath" -ForegroundColor Gray
            $foundRevit += @{
                Version = $version
                Path = $revitPath
                ApiPath = $apiPath
            }
        }
    }
}

if ($foundRevit.Count -eq 0) {
    Write-Host "❌ 未找到 Revit 安装" -ForegroundColor Red
    Write-Host "请确认 Revit 已安装在默认路径" -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "=== 推荐使用的 Revit 版本 ===" -ForegroundColor Cyan
    $latest = $foundRevit[0]
    Write-Host "Revit $($latest.Version)" -ForegroundColor Green
    Write-Host "API 路径: $($latest.ApiPath)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
