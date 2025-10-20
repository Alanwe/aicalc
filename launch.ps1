# AiCalc Studio - Quick Run (No Build)
# Launches the last built version after cleaning old processes

param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

# Kill any running instances to ensure we launch the latest version
Write-Host "Stopping any running instances..." -ForegroundColor Yellow
Stop-Process -Name "AiCalc.WinUI" -Force -ErrorAction SilentlyContinue

# Try the direct build output first (most common after dotnet build)
$exePath = "src\AiCalc.WinUI\bin\$Platform\$Configuration\net8.0-windows10.0.19041.0\AiCalc.WinUI.exe"

# Fallback to published version if available
if (-not (Test-Path $exePath)) {
    $exePath = "src\AiCalc.WinUI\bin\$Platform\$Configuration\net8.0-windows10.0.19041.0\win-$Platform\publish\AiCalc.WinUI.exe"
}

if (Test-Path $exePath) {
    Write-Host "Launching AiCalc Studio..." -ForegroundColor Green
    Write-Host "Path: $exePath" -ForegroundColor Gray
    Start-Process $exePath
} else {
    Write-Host "Application not built yet. Run .\run.ps1 first to build." -ForegroundColor Yellow
    Write-Host "Expected: src\AiCalc.WinUI\bin\$Platform\$Configuration\net8.0-windows10.0.19041.0\AiCalc.WinUI.exe" -ForegroundColor Gray
}
