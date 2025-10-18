# Quick Start - AiCalc WinUI

## ✅ What Works Now

The project is a **native Windows App SDK (WinUI 3)** application with **Phase 5 complete** and **Phase 6 in progress** - featuring AutoSave, CSV export/import, full UI polish, settings persistence, undo/redo, and formula syntax highlighting!

### Current Status:
```
✅ .NET 8.0 SDK installed
✅ Windows App SDK 1.4 configured  
✅ All business logic implemented (Models, Services, ViewModels, Converters)
✅ Phase 5 UI features complete (keyboard nav, context menus, themes, undo/redo)
✅ Phase 6 features partial (AutoSave, CSV export/import)
✅ Settings persistence working (window size, panels, theme)
✅ Project builds with 0 warnings, 0 errors
✅ All 59 tests passing
✅ Application runs smoothly on Windows
```

## 🚀 Commands

### Build the project:
```powershell
dotnet build AiCalc.sln
```

### Run the application:
```powershell
dotnet run --project src/AiCalc.WinUI/AiCalc.WinUI.csproj
# Or use the launch script:
.\launch.ps1
```

### Run tests:
```powershell
dotnet test tests/AiCalc.Tests/AiCalc.Tests.csproj
```

### Clean and rebuild:
```powershell
dotnet clean AiCalc.sln
dotnet build AiCalc.sln
```

## 📁 Project Structure

### Active Project (WinUI 3):
- **Location**: `src/AiCalc.WinUI/`
- **Project File**: `AiCalc.WinUI.csproj`
- **Entry Point**: `App.xaml.cs`
- **Main Window**: `MainWindow.xaml`

### Core Components:
- **Models/**: Data structures (CellAddress, CellChangeAction, UserPreferences, etc.)
- **Services/**: Business logic (FunctionRegistry, EvaluationEngine, UndoRedoManager, AI services)
- **ViewModels/**: MVVM layer (WorkbookViewModel, SheetViewModel, CellViewModel)
- **Converters/**: UI converters for WinUI bindings
- **Themes/**: Cell visual state themes (Light/Dark/High Contrast)

## 🎯 Phase 5 Features (Complete)

### Keyboard Shortcuts:
- **F9**: Recalculate all cells
- **F2**: Enter edit mode
- **Ctrl+Z**: Undo
- **Ctrl+Y**: Redo
- **Arrow Keys**: Navigate cells
- **Tab/Shift+Tab**: Move right/left
- **Enter/Shift+Enter**: Move down/up
- **Ctrl+Home/End**: Jump to first/last cell
- **Ctrl+Arrow**: Jump to data edge
- **Delete**: Clear cell

### Context Menu (Right-click):
- Cut, Copy, Paste, Clear Contents
- Insert Row Above/Below
- Insert Column Left/Right
- Delete Row/Column

### Settings Persistence:
- Window size and position
- Panel widths and visibility
- Theme preferences
- Recent workbooks (up to 10)
- Saved to: `%LocalAppData%\AiCalc\preferences.json`

### Undo/Redo:
- 50-action history
- Tracks value, formula, format changes
- Full command pattern implementation

### Formula Highlighting:
- Real-time tokenization
- Shows function and cell reference counts
- Supports sheet references (Sheet1!A1)

## 🎯 Phase 6 Features (In Progress)

### AutoSave:
- Timer-based automatic saving (1-60 minute intervals, default: 5 min)
- Dirty flag tracking (only saves when workbook has changed)
- Backup files: `filename_autosave.aicalc`
- Status notifications for save success/failure

### CSV Export/Import:
- **Export CSV**: Export current sheet to CSV format
- **Import CSV**: Import CSV as new sheet
- Proper CSV escaping (quotes, commas, newlines)
- UTF-8 encoding
- Robust parsing with quote handling
- UI buttons: 📤 Export CSV, 📥 Import CSV

### Usage:
```
Export: Click "📤 Export CSV" button
Import: Click "📥 Import CSV" button, select CSV file in file picker
```

## 🔧 What Was Fixed

### Source Code Issues (Fixed):
1. ✅ FunctionRegistry.cs - Guid formatting
2. ✅ FunctionRegistry.cs - JSON string escaping  
3. ✅ FunctionRunner.cs - Regex pattern escaping
4. ✅ CellAddress.cs - Integer division operator
5. ✅ App.xaml.cs - UnhandledExceptionEventArgs ambiguity
6. ✅ Converters - Visibility.Hidden → Visibility.Collapsed

### Architecture Change:
- **Before**: Uno Platform cross-platform project (had XAML compiler issues)
- **After**: Native Windows App SDK (WinUI 3) project ✅

## 📋 Next Development Steps

### Option 1: Use Simple UI (Current)
The app currently has a minimal placeholder UI. You can:
- Add basic controls to MainWindow.xaml
- Wire up ViewModels  
- Test business logic incrementally

### Option 2: Restore Full Spreadsheet UI
To get the complete spreadsheet interface:

1. Copy the backup XAML:
   ```powershell
   Copy-Item src/AiCalc/MainPage.xaml.bak src/AiCalc.WinUI/MainWindow.xaml
   ```

2. Update the XAML root element:
   - Change `x:Class="AiCalc.MainPage"` → `x:Class="AiCalc.MainWindow"`

3. Initialize WorkbookViewModel in MainWindow.xaml.cs

4. Build and run!

## 💡 Tips

### If you get build errors:
```powershell
# Clean everything
dotnet clean src/AiCalc.WinUI/AiCalc.WinUI.csproj
Remove-Item -Recurse -Force src/AiCalc.WinUI/obj, src/AiCalc.WinUI/bin

# Restore and rebuild
dotnet restore src/AiCalc.WinUI/AiCalc.WinUI.csproj
dotnet build src/AiCalc.WinUI/AiCalc.WinUI.csproj
```

### To check what's installed:
```powershell
dotnet --version          # Should show 8.0.200
dotnet --list-sdks        # Should include .NET 8.0
```

## 🎯 Core Features Available

All business logic is ready to use:

### Spreadsheet Functions:
- SUM, AVERAGE, COUNT
- CONCAT, UPPER, LOWER
- TEXT_TO_IMAGE (AI function)
- IMAGE_TO_CAPTION (AI function)
- DIRECTORY_TO_TABLE
- FILE_SIZE, FILE_EXTENSION
- And more in FunctionRegistry.cs

### Data Models:
- WorkbookDefinition with multiple sheets
- CellDefinition with formulas, values, automation modes
- Proper cell addressing (Sheet1!A1 format)
- JSON serialization (.aicalc files)

### MVVM Architecture:
- ObservableObject base classes
- ICommand implementations with CommunityToolkit
- Proper change notifications
- Ready for UI binding

## ✨ You're Ready to Build!

The foundation is solid. All the hard work of fixing syntax errors, resolving dependencies, and setting up the proper Windows App SDK environment is complete. Now you can focus on building features! 🚀
