# AiCalc Features Documentation

**Last Updated:** October 2025  
**Version:** 1.0  
**Status:** Active Development

---

## Overview

AiCalc is an AI-native spreadsheet application that combines traditional spreadsheet functionality with modern AI capabilities. This document tracks all features from the original specification and tasks.md, including their implementation and testing status.

---

## Legend

- ✅ **Implemented & Tested** - Feature is complete with automated tests
- 🟢 **Implemented** - Feature is complete but not yet tested
- 🟡 **Partially Implemented** - Feature is partially complete
- ⏳ **In Progress** - Currently being developed
- ⏭️ **Skipped** - Intentionally skipped due to technical limitations
- ❌ **Not Started** - Not yet implemented

---
# AiCalc Feature & Task Progress

**Last Updated:** 12 Oct 2025  
**Source of truth:** `tasks.md`

---

## How To Read This Document

Legend: ✅ Complete · 🟡 Partial/Needs follow-up · ❌ Not Started

References point to the files that implement (or miss) the corresponding work. When a requirement has notable gaps, they are listed explicitly so we can see what is left.

Current focus: **Phase 3 – Multi-Threading & Dependency Management.** Phases 1 and 4 still have open follow-ups, so we have not truly exited Phase 3 yet.

---

## Phase 1 – Core Infrastructure & Object-Oriented Cell System

- **Task 1 – Enhanced Cell Object Type System (✅)**  `Models/CellObjectType.cs`, `Models/CellObjects/*.cs`, `CellObjectFactory.cs`
   - Enum covers the requested object types; dedicated classes exist with validation/operations. Lacks dedicated tests for each cell object, but core behaviour is exercised indirectly via view models.
- **Task 2 – Type-Specific Function Registry (🟡)**  `Services/FunctionDescriptor.cs`, `Services/FunctionRegistry.cs`, `MainWindow.xaml.cs`
   - Function metadata and polymorphic parameters are implemented. UI still shows the full catalog; `GetFunctionsForTypes` is unused, so the registry is not yet cell-type aware.
- **Task 3 – Built-in Function Catalog (🟡)**  `Services/FunctionRegistry.cs`
   - Math/text/date basics exist. Table/file/directory helpers are mostly stubs (no TABLE_JOIN/TABLE_AGGREGATE logic, no FILE_WRITE/DIR_CREATE). Additional functions listed in `tasks.md` remain to be implemented and tested.

## Phase 2 – Cell Operations & Navigation

- **Task 4 – Excel-Like Keyboard Navigation (✅)**  `MainWindow.xaml.cs` (`MainWindow_KeyDown` et al.)
   - Full navigation surface implemented (arrows, Ctrl+Home/End, F2, F9, Delete, Enter behaviour toggle).
- **Task 5 – Cell Editing & Formula Intellisense (🟡)**  `MainWindow.xaml`, `MainWindow.xaml.cs`
   - Autocomplete, parameter hints, click-to-insert references, dependency highlighting in place. Missing items: syntax colouring, true range picker/colour coding, real-time validation beyond simple checks.
- **Task 6 – Value State vs Function State (❌)**
   - Cell inspector exposes value/formula fields but there is no explicit toggle, preview panes, or markdown/JSON editors described in the task.
- **Task 7 – Extraction, Spill & Insert Operations (🟡)**  `SheetViewModel.cs`, `ExtractFormulaDialog.cs`
   - Insert/delete rows & columns, extract formula dialog, basic spill support confirmed. Spill currently overwrites rather than shifting grid sections; reference reconciliation after structural changes still needs work.

## Phase 3 – Multi-Threading & Dependency Management

- **Task 8 – Dependency Graph (✅)**  `Services/DependencyGraph.cs`, tests in `tests/AiCalc.Tests/DependencyGraphTests.cs`
   - DAG with circular detection, range expansion, topological order implemented and covered by unit tests.
- **Task 9 – Multi-Threaded Evaluation (🟡)**  `Services/EvaluationEngine.cs`, `ViewModels/WorkbookViewModel.cs`, `SettingsDialog.xaml`
   - Batching/topological evaluation + settings wiring done. Outstanding items: progress feedback isn’t surfaced in UI, cancellation tokens aren’t exposed, per-service timeout overrides are not applied.
- **Task 10 – Cell Change Visualisation (🟡)**  `Models/CellVisualState.cs`, `Converters/CellVisualStateToBrushConverter.cs`, `App.xaml.cs`
   - Visual flashes, F9 shortcut and toolbar button, theme selection implemented. Manual-update state is never set, dependency chain highlight only appears during formula editing, and there is no recalculation progress overlay.

## Phase 4 – AI Functions & Service Integration

- **Task 11 – AI Service Configuration (🟡)**  `Models/WorkspaceConnection.cs`, `Services/AI/AIServicesRegistry.cs`, `ServiceConnectionDialog.xaml(.cs)`
   - Dialog exists but API keys are stored raw (no `CredentialService.Encrypt` usage) and connections are not registered with `App.AIServices`, so AI functions cannot currently locate a client.
- **Task 12 – AI Function Execution & Preview (❌)**
   - `FunctionRunner` routes AI calls, but the registry returns placeholder results (`[AI Processing Required]`) and no preview/regenerate flow is implemented. Lack of registered clients (Task 11) blocks real execution.
- **Task 13 – AI Functions + Cell Types (❌)**
   - No automatic suggestion/preview/cost estimation yet. AI functions are neither filtered by cell type nor surfaced in the UI.

## Phase 5 – Advanced UI/UX Features

- **Task 14 – Resizable/Collapsible Panels & Navigation (✅)**  `MainWindow.xaml`, `MainWindow.xaml.cs`
   - Keyboard navigation (Task 14A) and splitter-based panels with hide/show buttons (Task 14B) completed; state persisted in `WorkbookSettings`.
- **Task 15 – Rich Cell Editing Dialogs (❌)**
   - Only simple TextBox-based editors exist; markdown/json/code dialogs were never implemented due to WinUI issues.
- **Task 16 – Context Menu (✅)**  `MainWindow.xaml`, `MainWindow.xaml.cs`, dialogs under `src/AiCalc.WinUI`
   - Full context menu (cut/copy/paste, insert/delete, history, format, extract) delivered.

## Phase 6 – Data Sources & External Connections

- **Task 18 – Data Sources Menu (❌)**  Not started.
- **Task 19 – Cell-Level Data Binding (❌)**  Not started.

## Phase 7 – Python SDK & Scripting

- **Task 20 – Python SDK (🟡)**  `sdk/python/src/aicalc/client.py`, `Services/PipeServer.cs`
   - Named pipe IPC works; CRUD/formula/function helpers exist. Missing items: environment discovery, pandas integration, event subscriptions, automated tests, and the legacy `python-sdk/` package is out-of-date.
- **Task 21 – Python Function Discovery (❌)**  No decorator scan / registry integration yet.
- **Task 22 – Cloud Scripts (❌)**  Not started. - Dont implement yet 

## Phase 8 – Advanced Features & Polish

- **Task 23 – Cell Inspector Tabs (❌)**  Inspector still a single column, no tabbed layout.
- **Task 24 – Automation System (🟡)**  Basic automation modes exist (`CellViewModel.AutomationMode`), but there are no triggers/flows/UI.
- **Task 25 – File Format & Persistence (🟡)**  `WorkbookViewModel.SaveAsync/LoadAsync` handle JSON. Binary, autosave, CSV/XLSX import/export remain outstanding.
- **Task 26 – Row/Column Operations (🟡)**  Insert/delete implemented; hide/show, resizing, freezing, grouping pending.
- **Task 27 – Selection & Range Operations (❌)**  No multi-select, fill, or formula propagation yet.
- **Task 28 – Themes & Customisation (🟡)**  Application & cell-state themes selectable; there is no custom colour editor beyond presets.
- **Task 29 – Community Functions (❌)**
- **Task 30 – Plans & Monetisation (❌)**

## Phase 9 – Testing, Documentation & Deployment

- **Task 31 – Unit Testing (🟡)**  `tests/AiCalc.Tests`
   - 36 `[Fact]` specs (~40 total tests with `[Theory]`) cover models/settings/dependency graph. Function registry, evaluation engine, AI integrations, and UI logic remain untested.
- **Task 32 – Integration Testing (❌)**  No automated end-to-end or UI automation.
- **Task 33 – Documentation (🟡)**  README, tasks.md, STATUS.md exist; user guide, API reference, SDK docs pending (this file now reflects actual progress).
- **Task 34 – Performance Optimisation (❌)**
- **Task 35 – Deployment & Distribution (❌)**

---

## Current Test & Build Snapshot

- Unit tests: ~40 (xUnit) covering models + dependency graph only. No coverage for evaluation engine, AI glue, or UI logic.
- Build: `dotnet build AiCalc.sln` succeeds (see `STATUS.md`).
- Manual testing: Spreadsheet UI renders, keyboard shortcuts function, formula evaluation works for synchronous built-ins; AI functions currently return placeholders due to missing client registration.

---

## Key Gaps Before Advancing Beyond Phase 3

1. Close remaining Phase 1/2 items: finish type-aware function surfacing (Task 2) and value/formula state UX (Task 6).
2. Expose evaluation progress & cancellation, and add real tests for `EvaluationEngine` (Task 9).
3. Wire AI connection storage to `App.AIServices`, encrypt API keys, and deliver real AI call previews (Tasks 11–13).

Once those are in place we can legitimately move into Phase 4 feature work.

---

## Quick Reference – Useful Files

- Spreadsheet UI & behaviour: `src/AiCalc.WinUI/MainWindow.xaml`, `MainWindow.xaml.cs`
- View models & models: `src/AiCalc.WinUI/ViewModels`, `src/AiCalc.WinUI/Models`
- Function execution stack: `src/AiCalc.WinUI/Services/FunctionRegistry.cs`, `FunctionRunner.cs`
- Evaluation pipeline: `src/AiCalc.WinUI/Services/DependencyGraph.cs`, `EvaluationEngine.cs`
- AI service plumbing: `src/AiCalc.WinUI/Services/AI/*`
- Python IPC bridge: `src/AiCalc.WinUI/Services/PipeServer.cs`, `sdk/python/src/aicalc/client.py`
- Unit tests: `tests/AiCalc.Tests/*`

---

Maintained by: GitHub Copilot (coding agent) + AiCalc team.
**Description:** Standard spreadsheet navigation with keyboard shortcuts.
