# AiCalc Studio - Quick Launch Script
# This script builds and runs the application

param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

Write-Host "Building AiCalc Studio..." -ForegroundColor Cyan
dotnet publish src/AiCalc.WinUI/AiCalc.WinUI.csproj -c $Configuration -r win-$Platform --self-contained /p:Platform=$Platform

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nLaunching AiCalc Studio..." -ForegroundColor Green
    $exePath = "src\AiCalc.WinUI\bin\$Platform\$Configuration\net8.0-windows10.0.19041.0\win-$Platform\publish\AiCalc.WinUI.exe"
    
    if (Test-Path $exePath) {
        Start-Process $exePath
        Write-Host "Application launched successfully!" -ForegroundColor Green
    } else {
        Write-Host "Error: Executable not found at $exePath" -ForegroundColor Red
    }
} else {
    Write-Host "`nBuild failed!" -ForegroundColor Red
}
