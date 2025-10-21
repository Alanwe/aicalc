# Debug launch script - runs app with console output
param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

Write-Host "Building AiCalc for debugging..." -ForegroundColor Cyan
dotnet build src/AiCalc.WinUI/AiCalc.WinUI.csproj -c $Configuration /p:Platform=$Platform

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nStarting AiCalc with debug output..." -ForegroundColor Green
    $exePath = "src\AiCalc.WinUI\bin\$Platform\$Configuration\net8.0-windows10.0.19041.0\AiCalc.WinUI.exe"
    
    if (Test-Path $exePath) {
        Write-Host "Debug output will appear below. Press Ctrl+C to stop." -ForegroundColor Yellow
        Write-Host "================================================" -ForegroundColor Cyan
        
        try {
            $process = Start-Process -FilePath $exePath -Wait -PassThru
            Write-Host "`nApp exited with code: $($process.ExitCode)" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error running app: $_" -ForegroundColor Red
        }
        
        Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } else {
        Write-Host "Error: Executable not found at $exePath" -ForegroundColor Red
    }
} else {
    Write-Host "`nBuild failed!" -ForegroundColor Red
}
