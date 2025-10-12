# Installing and running AiCalc Studio

This guide walks through preparing your environment, restoring the project, and running the WinUI 3 application on Windows.

## Prerequisites

1. **Windows 10 version 1809 or later** (Windows 11 recommended for best experience)
2. **Visual Studio 2022** (17.0 or later) with the following workloads:
   - .NET Desktop Development
   - Universal Windows Platform development (for Windows App SDK)
3. **.NET 8 SDK** - [Download from Microsoft](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
4. **Windows App SDK 1.4 or later** - Usually installed with Visual Studio workloads

### Verifying Prerequisites

Check that .NET 8 is installed:
```bash
dotnet --version
```

You should see version 8.0.x or later.

## Cloning the repository

```bash
git clone https://github.com/Alanwe/aicalc.git
cd aicalc
```

## Restoring dependencies

```bash
dotnet restore
```

This will restore all NuGet packages including:
- Microsoft.WindowsAppSDK 1.4.231219000
- CommunityToolkit.Mvvm 8.2.1
- System.Text.Json 8.0.5
- System.Security.Cryptography.ProtectedData 8.0.0

## Building the project

Build the WinUI 3 project:

```bash
dotnet build src/AiCalc.WinUI/AiCalc.WinUI.csproj
```

Or build the entire solution:

```bash
dotnet build AiCalc.sln
```

Expected output:
```
Build succeeded.
    1 Warning(s)  (NETSDK1206 - can be ignored)
    0 Error(s)
```

## Running the application

### Option 1: Using dotnet CLI

```bash
dotnet run --project src/AiCalc.WinUI/AiCalc.WinUI.csproj
```

### Option 2: Using Visual Studio

1. Open `AiCalc.sln` in Visual Studio 2022
2. Set `AiCalc.WinUI` as the startup project (if not already)
3. Press **F5** to run or **Ctrl+F5** to run without debugging

### Option 3: Using the PowerShell scripts

```powershell
# Build the project
.\run.ps1

# Or use the launch script
.\launch.ps1
```

## Running Tests

Run all unit tests:

```bash
dotnet test tests/AiCalc.Tests/AiCalc.Tests.csproj
```

## Troubleshooting

### Build Errors

**Error: NETSDK1100 "To build a project targeting Windows..."**
- This means you're trying to build on a non-Windows OS
- WinUI 3 requires Windows for building and running
- Solution: Use a Windows machine or VM

**Error: Cannot find Windows SDK**
- Install the Windows 10 SDK through Visual Studio Installer
- Or download directly from [Microsoft](https://developer.microsoft.com/windows/downloads/windows-sdk/)

**Error: XamlCompiler.exe exited with code 1**
- This is a known issue with certain XAML patterns in WinUI 3
- The current codebase avoids problematic patterns
- If you see this, you may have added GridSplitter or complex ContentDialog layouts

### Runtime Warnings

**Warning: NETSDK1206 (Runtime identifier warning)**
- This warning is informational and can be ignored
- The application builds and runs correctly despite this warning
- Related to Windows App SDK RID naming conventions

### Application Won't Start

1. Ensure you're running on Windows 10 1809 or later
2. Verify Windows App SDK is installed:
   - Check in Visual Studio Installer under Individual Components
   - Look for "Windows App SDK C# Templates"
3. Try rebuilding:
   ```bash
   dotnet clean
   dotnet restore
   dotnet build
   ```

## Development Workflow

1. Open `AiCalc.sln` in Visual Studio 2022
2. Make code changes to files in `src/AiCalc.WinUI/`
3. Press **F5** to build and run
4. Use **Hot Reload** (Alt+F10) for quick UI changes
5. Run tests with **Test Explorer** (Ctrl+E, T)

## Project Structure

```
AiCalc.sln                      # Visual Studio solution
src/
  AiCalc.WinUI/                 # Main WinUI 3 application
    AiCalc.WinUI.csproj         # Project file
    App.xaml / .cs              # Application entry point
    MainWindow.xaml / .cs       # Main spreadsheet UI
    SettingsDialog.xaml / .cs   # Settings dialog
    Models/                     # Business models
    ViewModels/                 # MVVM view models
    Services/                   # Business logic & AI clients
    Themes/                     # UI theme resources
    Converters/                 # XAML value converters
tests/
  AiCalc.Tests/                 # Unit tests
    AiCalc.Tests.csproj         # Test project
```

## Next Steps

- See [README.md](README.md) for feature overview
- See [STATUS.md](STATUS.md) for current development status
- See [QUICKSTART.md](QUICKSTART.md) for quick command reference
- See `docs/` folder for detailed phase documentation

## AI Service Configuration

To use AI features:

1. Launch the application
2. Click **Settings** button (or press F9 then click Settings)
3. Go to **AI Services** tab
4. Add a connection:
   - **Azure OpenAI**: Provide endpoint URL and API key
   - **Ollama**: Provide local endpoint (e.g., http://localhost:11434)
5. Configure models for Text, Vision, and Image generation
6. Click **Test Connection** to verify
7. Set as default connection

For more details, see `docs/Phase4_COMPLETE.md`.

## Getting Help

- **Build Issues**: Check troubleshooting section above
- **Feature Questions**: See documentation in `docs/` folder
- **Bug Reports**: File an issue on GitHub
