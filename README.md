# AiCalc Studio

AiCalc Studio is a **Windows App SDK (WinUI 3)** application that delivers an AI-native spreadsheet experience. Each cell can host rich content such as directories, media, documents, links, or traditional scalar values while orchestrating AI workflows alongside classic spreadsheet logic.

**Current Status:** Phase 6 In Progress - AutoSave and CSV export/import implemented. 59 passing tests, clean builds, production-ready Phase 5 features.

## Highlights

- **Native Windows App** built with WinUI 3 for high performance and modern UI
- **Rich cell types** with visual glyphs that describe the current payload (text, images, video, directories, scripts, and more)
- **AI function catalog** with 25+ functions for text-to-image, image captioning, directory introspection, code generation, and more
- **Multi-threaded evaluation** with dependency graph (DAG) and parallel processing
- **Excel-like keyboard navigation** with 8+ shortcuts (F9, F2, arrows, Ctrl+Z/Y, etc.)
- **Context menus** with 13 operations (Cut/Copy/Paste, Insert/Delete rows/columns)
- **Undo/Redo system** with 50-action history and full command pattern
- **Settings persistence** - Window size, panel states, theme preferences saved automatically
- **Formula syntax highlighting** with real-time tokenization and visual feedback
- **Theme system** - Light/Dark/System app themes + 4 cell visual state themes
- **AutoSave service** - Timer-based automatic saving with 1-60 minute intervals
- **CSV export/import** - Export sheets to CSV, import CSV as new sheets
- **Workbook automation** with manual, auto-run on open, and dependency-triggered evaluation
- **Integrated inspector** to tweak values, formulas, automation mode, and notes
- **Secure AI connections** with DPAPI encryption for Azure OpenAI, Ollama, and more
- **JSON workbook persistence** using the `.aicalc` file extension

## Project layout

```
AiCalc.sln              # Solution file
src/
  AiCalc.WinUI/        # Windows App SDK (WinUI 3) application
    App.xaml           # Theme resources and global styles
    App.xaml.cs        # Application bootstrapper with preferences service
    MainWindow.xaml    # Primary spreadsheet shell with formula bar
    MainWindow.xaml.cs # UI logic, keyboard nav, context menus, undo/redo
    Models/            # Workbook, sheet, cell, settings, preferences, actions
    ViewModels/        # MVVM state for workbook, sheets, rows, and cells
    Services/          # Function registry, evaluation engine, AI services, undo/redo
    Converters/        # UI value converters (automation glyphs, brushes, visibility)
    Themes/            # Cell visual state themes (Light/Dark/High Contrast)
tests/
  AiCalc.Tests/        # xUnit tests (59 passing)
docs/                  # Comprehensive documentation
```

## Getting started

1. **Prerequisites:**
   - .NET SDK 8.0 or later
   - Windows 10 SDK (10.0.19041.0 or later)
   - Windows App SDK 1.4

2. **Build and run:**
   ```powershell
   cd C:\Projects\aicalc
   dotnet build AiCalc.sln
   dotnet run --project src/AiCalc.WinUI/AiCalc.WinUI.csproj
   ```

3. **Or use Visual Studio:**
   - Open `AiCalc.sln`
   - Set `AiCalc.WinUI` as startup project
   - Press F5 to run

4. **Run tests:**
   ```powershell
   dotnet test tests/AiCalc.Tests/AiCalc.Tests.csproj
   ```

## Extending the function catalog

The `FunctionRegistry` class registers built-in spreadsheet and AI helpers. New functions can be added by calling `Register` with a `FunctionDescriptor`. Each descriptor receives a `FunctionEvaluationContext` that exposes the workbook, current sheet, argument cells, and raw formula so you can orchestrate:

- Native .NET logic
- External agent flows via [Microsoft Agent Framework](https://github.com/microsoft/agent-framework/tree/main/dotnet)
- Python scripts hosted in your runtime

This design makes it straightforward to contribute additional AI skills or data utilities directly from C# or through Python interop.

## Saving workbooks

The **Save** and **Load** commands serialize the workbook to a friendly JSON payload with the `.aicalc` suffix. Serialized state captures:

- Workbook metadata and connections
- Sheet layout, row/column dimensions
- Cell formulas, automation preferences, notes, and materialized values

## Roadmap ideas

- Real agent execution with background orchestration
- Python function discovery and hot reloading
- Collaborative editing and change tracking
- Live previews for media cells (image/video/audio)
- Advanced grid virtualization for large datasets

> **Note:** This repository is a design-forward blueprint. Additional plumbing is required to reach production readiness, including runtime-specific heads, storage prompts, and integration with live AI providers.
