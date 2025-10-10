# AiCalc Studio - Quick Run (No Build)
# Launches the last built version

param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

$exePath = "src\AiCalc.WinUI\bin\$Platform\$Configuration\net8.0-windows10.0.19041.0\win-$Platform\publish\AiCalc.WinUI.exe"

if (Test-Path $exePath) {
    Write-Host "Launching AiCalc Studio..." -ForegroundColor Green
    Start-Process $exePath
} else {
    Write-Host "Application not built yet. Run .\run.ps1 first to build and publish." -ForegroundColor Yellow
    Write-Host "Looking for: $exePath" -ForegroundColor Gray
}
