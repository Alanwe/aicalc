# AiCalc Studio

AiCalc Studio is a concept UNO Platform application that delivers an AI-native spreadsheet experience. Each cell can host rich content such as directories, media, documents, links, or traditional scalar values while orchestrating AI workflows alongside classic spreadsheet logic.

## Highlights

- **Cross-platform UNO UI** targeting Windows, WebAssembly, and mobile platforms.
- **Rich cell types** with visual glyphs that describe the current payload (text, images, video, directories, scripts, and more).
- **AI function catalog** with starter functions for text-to-image, image captioning, directory introspection, and familiar numeric or string utilities.
- **Workbook automation** that supports manual execution, auto-run on open, or dependency-triggered evaluation.
- **Integrated inspector** to tweak values, formulas, automation mode, and notes while previewing output instantly.
- **Connection settings** for registering local runtimes, Ollama, Azure OpenAI, or additional AI backends.
- **JSON workbook persistence** using the `.aicalc` file extension for saving and loading multi-sheet projects.

## Project layout

```
AiCalc.sln              # Solution file
src/
  AiCalc/              # UNO Platform single-project app
    App.xaml           # Theme resources and global styles
    App.xaml.cs        # Application bootstrapper
    MainPage.xaml      # Primary shell with navigation and spreadsheet surface
    Models/            # Workbook, sheet, cell, and settings models
    ViewModels/        # MVVM state for workbook, sheets, rows, and cells
    Services/          # Function registry and evaluation engine
    Converters/        # UI value converters (automation glyphs, brushes, visibility)
```

## Getting started

1. Install the [UNO Platform prerequisites](https://platform.uno/docs/articles/get-started.html).
2. Restore and build the solution:
   ```bash
   dotnet restore
   dotnet build
   ```
3. Run the desired head (Windows, WebAssembly, Android, iOS, or macOS) from Visual Studio or with `dotnet run -f net7.0-windows`.

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
