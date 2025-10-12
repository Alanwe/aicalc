# AiCalc Studio

AiCalc Studio is a Windows App SDK (WinUI 3) application that delivers an AI-native spreadsheet experience. Each cell can host rich content such as directories, media, documents, links, or traditional scalar values while orchestrating AI workflows alongside classic spreadsheet logic.

## Highlights

- **Native Windows UI** powered by WinUI 3 and Windows App SDK
- **Rich cell types** with visual glyphs that describe the current payload (text, images, video, directories, scripts, and more)
- **AI function catalog** with starter functions for text-to-image, image captioning, directory introspection, and familiar numeric or string utilities
- **Workbook automation** that supports manual execution, auto-run on open, or dependency-triggered evaluation
- **Integrated inspector** to tweak values, formulas, automation mode, and notes while previewing output instantly
- **Connection settings** for registering local runtimes, Ollama, Azure OpenAI, or additional AI backends
- **JSON workbook persistence** using the `.aicalc` file extension for saving and loading multi-sheet projects
- **Keyboard navigation** with Excel-like shortcuts (F2, F9, Arrow keys, Ctrl+Home/End, Tab, Enter, etc.)
- **Context menus** for Cut/Copy/Paste, Insert/Delete rows/columns
- **Theme system** supporting Light/Dark/System app themes and custom cell state colors

## Project layout

```
AiCalc.sln                   # Solution file
src/
  AiCalc.WinUI/             # WinUI 3 application
    App.xaml                # Theme resources and global styles
    App.xaml.cs             # Application bootstrapper with theme logic
    MainWindow.xaml         # Primary spreadsheet interface
    MainWindow.xaml.cs      # Main window code-behind with keyboard/context menus
    SettingsDialog.xaml     # Settings UI for AI services and themes
    Themes/                 # Theme resource dictionaries
    Models/                 # Workbook, sheet, cell, and settings models
    ViewModels/             # MVVM state for workbook, sheets, rows, and cells
    Services/               # Function registry, evaluation engine, AI clients
    Converters/             # UI value converters (automation glyphs, brushes, visibility)
tests/
  AiCalc.Tests/             # Unit tests for core functionality
```

## Getting started

1. Prerequisites:
   - Windows 10 version 1809 or later
   - .NET 8.0 SDK
   - Windows App SDK 1.4 or later
   - Visual Studio 2022 (recommended) with Windows App SDK workload
   
2. Restore and build the solution:
   ```bash
   dotnet restore
   dotnet build
   ```
   
3. Run the application:
   ```bash
   dotnet run --project src/AiCalc.WinUI/AiCalc.WinUI.csproj
   ```
   
   Or open `AiCalc.sln` in Visual Studio 2022 and press F5.

## Extending the function catalog

The `FunctionRegistry` class registers built-in spreadsheet and AI helpers. New functions can be added by calling `Register` with a `FunctionDescriptor`. Each descriptor receives a `FunctionEvaluationContext` that exposes the workbook, current sheet, argument cells, and raw formula so you can orchestrate:

- Native .NET logic
- AI service integration via Azure OpenAI or Ollama
- Custom data processing workflows

This design makes it straightforward to contribute additional AI skills or data utilities directly from C#.

## AI Service Integration

AiCalc supports multiple AI providers:

- **Azure OpenAI**: GPT-4 for text completion, GPT-4-Vision for image captioning, DALL-E 3 for image generation
- **Ollama**: Local AI models (llama2, mistral, codellama, LLaVA for vision)
- Extensible architecture for additional providers

Configure AI services through the Settings dialog (accessed via toolbar or F9 settings).

## Keyboard Shortcuts

AiCalc includes Excel-like keyboard navigation:

- **F9**: Recalculate all formulas
- **F2**: Edit current cell
- **Arrow Keys**: Navigate between cells
- **Tab/Shift+Tab**: Move right/left
- **Enter/Shift+Enter**: Move down/up
- **Ctrl+Home/End**: Jump to A1 or last cell
- **Ctrl+Arrow**: Jump to data edge
- **Page Up/Down**: Scroll by pages
- **Delete**: Clear cell contents

## Context Menus

Right-click on cells for quick actions:
- Cut, Copy, Paste
- Clear Contents
- Insert Row Above/Below
- Insert Column Left/Right
- Delete Row/Column

## Saving workbooks

The **Save** and **Load** commands serialize the workbook to a friendly JSON payload with the `.aicalc` suffix. Serialized state captures:

- Workbook metadata and connections
- Sheet layout, row/column dimensions
- Cell formulas, automation preferences, notes, and materialized values

## Roadmap ideas

- Advanced formula bar with autocomplete and syntax highlighting
- Settings persistence and undo/redo functionality
- Enhanced cell editing dialogs for Markdown, JSON, and images
- Python function discovery and hot reloading
- Collaborative editing and change tracking
- Live previews for media cells (image/video/audio)
- Advanced grid virtualization for large datasets
- Support for additional AI providers

## Current Development Status

### Phase 4: AI Functions & Service Integration ✅ COMPLETE (100%)
- 9 AI functions fully implemented
- Secure credential storage with Windows DPAPI
- Multi-provider support (Azure OpenAI, Ollama)
- Token tracking and usage statistics
- Enhanced settings UI with multi-model configuration

### Phase 5: UI Polish & Enhancements ✅ PARTIALLY COMPLETE (60%)
- ✅ Keyboard navigation with 8+ shortcuts
- ✅ Context menus with 13 operations
- ✅ Theme system (Light/Dark/System + cell state themes)
- ⏭️ Resizable panels (skipped due to WinUI 3 XAML compiler limitations)
- ⏭️ Rich cell editing dialogs (skipped due to WinUI 3 XAML compiler limitations)

See `docs/Phase4_COMPLETE.md` and `docs/Phase5_Summary.md` for detailed documentation.

> **Note:** This is an actively developed project. The core spreadsheet functionality, AI integration, and UI features are functional and ready for use.
