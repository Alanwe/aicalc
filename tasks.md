# AiCalc Implementation Tasks

## Overview
This document outlines the implementation tasks to transform AiCalc from its current state into a fully-featured AI-native spreadsheet application. Tasks are organized in logical implementation order, with dependencies clearly marked.

---

## Phase 1: Core Infrastructure & Object-Oriented Cell System

### Task 1: Enhance Cell Object Type System
**Goal**: Implement robust class-based cell type system where each cell type has specific functions that it is able to be used with.
**Dependencies**: None
**Details**:
- Extend `CellObjectType` enum with all types from ideas.md (PDF, PDFPage, Code-Python, Code-CSS, etc.)
- Create base `CellObject` interface/abstract class
- Implement specific cell object classes: `NumberCell`, `StringCell`, `ImageCell`, `DirectoryCell`, `FileCell`, `TableCell`, `VideoCell`, `PDFCell`, `PDFPageCell`, `MarkdownCell`, `JsonCell`, `XmlCell`, `CodeCell` (with language variants)
- Each class should have properties and methods relevant to its type
- Add validation and type-specific behaviors

**Questions**:
- Do you want PDF handling to be file-path based or byte-array based?
We should use off the shelf functions for PDF operations, ideally c# libraries or we could use Python Packages. Normally the input will be a reference to a file. I dont want too much unstructed data in Memory as we will bloat 
- Should code cells have syntax highlighting metadata stored?
Not yet
- What specific operations should each cell type support initially?
Examples could be PDFtoPageText(split into pages), PDFChartImage(Extract charts as Chart Image)

---

### Task 2: Implement Type-Specific Function Registry
**Goal**: Functions should only be available for relevant cell types
**Dependencies**: Task 1
**Details**:
- Extend `FunctionDescriptor` to include applicable `CellObjectType[]`
- Modify `FunctionRegistry` to filter functions by cell type
- Update UI to show only relevant functions for selected cell type
- Add function categories (Built-in, AI, Community/Contrib)
- Create function discovery mechanism based on cell object class

**Questions**:
- Should functions be able to accept multiple input types (polymorphic)?
Yes, Functions should be able to take multiple class objects as parameters, but each function parameter position should force acceptable classes
- Do you want a visual categorization in the Functions panel?
Yes

---

### Task 2: Implement Type-Specific Function Registry
**Goal**: Expand built-in function catalog
**Dependencies**: Task 2
**Details**:
- Add missing math functions: AVERAGE, COUNT, MIN, MAX, MEDIAN, ROUND
- Add text functions: UPPER, LOWER, TRIM, SPLIT, REPLACE, LEN
- Add date/time functions: NOW, TODAY, DATE, TIME, DATEADD, DATEDIFF
- Add file functions: FILE_SIZE, FILE_EXTENSION, FILE_READ, FILE_WRITE
- Add directory functions: DIR_LIST, DIR_SIZE, DIR_CREATE
- Add table functions: TABLE_FILTER, TABLE_SORT, TABLE_JOIN, TABLE_AGGREGATE
- Categorize all functions (Built-in vs AI vs Contrib)

---

## Phase 2: Cell Operations & Navigation

### Task 4: Excel-Like Keyboard Navigation
**Goal**: Implement standard spreadsheet navigation
**Dependencies**: None (can run parallel with Phase 1)
**Details**:
- Arrow keys: Move selection Up/Down/Left/Right
- Tab: Move right, Shift+Tab: Move left
- Enter: Save cell and move down
- Shift+Enter: Save cell and move up
- Ctrl+Home: Go to A1
- Ctrl+End: Go to last used cell
- Page Up/Down: Scroll viewport
- Ctrl+Arrow: Jump to edge of data region
- F2: Edit mode toggle

**Questions**:
- Should Enter behavior be configurable (move down vs stay)?
Yes that would be nice, perhaps we can build a settings configurations tab in the menu

---

### Task 5: Cell Editing & Formula Intellisense
**Goal**: Rich in-cell editing experience
**Dependencies**: Task 2, Task 4
**Details**:
- When typing "=" show formula mode indicator
- Implement function name autocomplete dropdown
- When typing function and opening "(", show parameter hints
- Implement cell/range picker (click cells to add to formula)
- Show cell references with colored highlights
- Add formula validation (real-time syntax checking)
- Support range notation (A1:A10)

**Questions**:
- Should we support Excel-style named ranges?
Yes the functions should accept ranges which would be an array/matrix of class objects.
- Do you want syntax highlighting in formula bar?
Not sure I understand, I will take your recommendation

---

### Task 6: Cell Value State vs Function State
**Goal**: Toggle between raw value editing and formula evaluation
**Dependencies**: Task 5
**Details**:
- Add "Value State" button/toggle to Cell Inspector
- In Value State: Show rich text editor for Markdown, JSON/XML formatted editor
- Add preview pane for Markdown cells
- Add JSON/XML validator and formatter
- Support manual value override (disable formula temporarily)
- Add visual indicator when cell is in value vs formula state

**Questions**:
- Should value state lock the cell from formula evaluation? Yes
- Do you want a side-by-side preview for Markdown? Yes please

---

### Task 7: Cell Extraction, Spill & Insert Operations
**Goal**: Advanced cell manipulation features
**Dependencies**: Task 6
**Details**:
- **Extract Function**: Right-click menu to extract formula to new cell
- **Spill/Overwrite**: Array formulas that automatically fill multiple cells
- **Insert**: Add cells with shift down/right options
- **Displace**: Insert with graceful reference updating
- **Cancel**: Undo recent operation
- Update all cell references when rows/columns are inserted/deleted
- Add warning dialog for overwrite operations

**Questions**:
- Should spill operations be automatic (like Excel dynamic arrays)?
The Spill should shift columns impacted to the right and impacted rows down. The whole column and row should shift the number of cells
- What should happen to references when cells are deleted?
The cells which reference the deleted should be flagged as error and the value set to N/A like excel

---

## Phase 3: Multi-Threading & Dependency Management

### Task 8: Dependency Graph (DAG) Implementation
**Goal**: Track cell dependencies for efficient evaluation
**Dependencies**: None (can run parallel)
**Status**: ✅ COMPLETE
**Details**:
- ✅ Implement Directed Acyclic Graph (DAG) for cell relationships
- ✅ Parse formulas to extract cell references
- ✅ Build dependency tree (which cells depend on which)
- ✅ Detect circular reference loops (show error)
- ✅ Detect duplicate calculations
- ✅ Use linked list structure for efficient traversal

**Implementation**: See `src/AiCalc.WinUI/Services/DependencyGraph.cs`

**Questions**:
- Should circular references be blocked or allowed with iteration limit? An Error should be flagged ✅ DONE
- How many levels of dependency should be visualized in UI?
The dependency shouldnt normally be visualised, this should be managed quietly as Excel does it. ✅ DONE

---

### Task 9: Multi-Threaded Cell Evaluation
**Goal**: Parallel formula evaluation
**Dependencies**: Task 8
**Status**: ⚠️ 70% COMPLETE (Core done, Settings UI pending)
**Details**:
- ✅ Implement topological sort of dependency graph
- ✅ Evaluate independent cells in parallel threads
- ✅ Use Task Parallel Library (TPL) for async evaluation
- ✅ Show progress indicator for large workbooks
- ✅ Implement cancellation tokens for long-running operations
- ✅ Add timeout configuration for AI functions (100s default)
- ✅ Queue system for batch evaluation
- ⏳ Settings UI for thread count configuration (PENDING)
- ⏳ Per-service timeout configuration (PENDING)

**Implementation**: See `src/AiCalc.WinUI/Services/EvaluationEngine.cs`

**Remaining Work** (~5-7 hours):
1. Create/extend SettingsDialog.xaml with Performance section
2. Add slider for MaxDegreeOfParallelism (1-32, default: CPU count)
3. Add input for DefaultTimeoutSeconds (10-300, default: 100)
4. Save settings to WorkbookSettings
5. Wire up settings to EvaluationEngine

**Questions**:
- What should be the default timeout for AI functions?
We should default to 100 seconds, this should be configurable a the AI Service definition ✅ DONE
- Should users be able to configure thread count?
This would be good as an option in the settings menu ⏳ PENDING

---

### Task 10: Cell Change Visualization
**Goal**: Visual feedback for cell updates
**Dependencies**: Task 9
**Status**: ⚠️ 70% COMPLETE (Visual states done, F9 & themes pending)
**Details**:
- ✅ When cell value changes: Flash green border for 2 seconds
- ✅ When cell is stale (needs recalc): Show blue border
- ✅ Manual update mode: Show orange indicator
- ✅ Add update timestamp to cell metadata
- ✅ Show "calculating..." spinner during evaluation
- ✅ Highlight cells in dependency chain when selected
- ⏳ F9 keyboard shortcut for recalculate all (PENDING)
- ⏳ Recalculate All button in toolbar (PENDING)
- ⏳ Theme system with customizable colors (PENDING)

**Implementation**: See:
- `src/AiCalc.WinUI/Models/CellVisualState.cs`
- `src/AiCalc.WinUI/Converters/CellVisualStateToBrushConverter.cs`
- `src/AiCalc.WinUI/ViewModels/CellViewModel.cs` (MarkAsStale, MarkAsCalculating, MarkAsUpdated methods)

**Remaining Work** (~7-9 hours):
1. **F9 Recalculation** (~2-3 hours):
   - Add KeyboardAccelerator in MainWindow.xaml
   - Create RecalculateAllCommand in SheetViewModel
   - Filter out Manual automation mode cells
   - Show progress during recalc

2. **Recalculate All Button** (~30 minutes):
   - Add button to toolbar
   - Bind to same command as F9
   - Tooltip: "Recalculate All (F9)"

3. **Theme System** (~4-5 hours):
   - Create theme resource dictionaries (Light/Dark/High Contrast)
   - Modify converter to use theme resources
   - Add theme selector to Settings
   - Support custom color overrides

**Questions**:
- Should color scheme be customizable/theme-aware?
Yes please. Perhaps some set themes which can be customised further? ⏳ PENDING
- Do you want a "recalculate all" button?
Yes and also this should be available as F9 like excel. However this should not apply to cells flags to manual which should be excluded. ⏳ PENDING

---

## Phase 4: AI Functions & Service Integration

### Task 11: AI Function Configuration System
**Goal**: Secure, configurable AI service connections
**Dependencies**: Task 2
**Details**:
- Extend `WorkspaceConnection` model with API key encryption
- Use Windows DPAPI for secure local storage of credentials
- Implement connection testing/validation
- Add connection selector in Cell Inspector
- Store AI function preferences per cell type
- Add token usage tracking and cost estimation
- Support multiple providers: Azure OpenAI, Ollama, OpenAI, Anthropic, etc.

**Questions**:
- Should API keys be stored per-workbook or globally?
For now lets go global 
- Do you want usage/cost reporting dashboard?
Yes, It would need to be usage for Basic plan as we wont know the actual costs
Some usage charts per service and perhaps average Tokens and avg latency would be useful

---

### Task 12: AI Function Execution & Preview
**Goal**: Execute AI functions with preview
**Dependencies**: Task 11
**Details**:
- Implement actual AI service calls (Azure OpenAI, Ollama)
- Add preview mode: show AI response before committing to cell
- Add regeneration button (try again with same prompt)
- Support streaming responses for long-running operations
- Add prompt template system
- Cache AI responses to avoid duplicate calls
- Add "AI Assistant" panel for multi-step operations

**Questions**:
- Should there be a default model per function type?
The AI Service should configure the model deployment, this is then linked to the functions so we can look it up
- Do you want a history of AI generations for each cell?
Can we have a log mode in settings, default is no

---

### Task 13: Tie AI Functions to Classes & Preview
**Goal**: Smart AI function routing based on cell type
**Dependencies**: Task 12, Task 2
**Details**:
- Map AI functions to specific cell object types
- Auto-suggest AI functions based on cell content
- Preview AI function output before execution
- Show estimated tokens/cost before running
- Add "batch AI" feature to run function on range
- Implement fallback/error handling for AI failures

---

## Phase 5: Advanced UI/UX Features - ✅ COMPLETE

### Task 14A: Keyboard Navigation - ✅ COMPLETE
**Goal**: Excel-like keyboard shortcuts
**Status**: ✅ Fully implemented
**Details**:
- ✅ 8+ keyboard shortcuts (F9, F2, arrows, Tab, Enter, Ctrl+Home/End, Ctrl+Arrow, Page Up/Down, Delete)
- ✅ Excel-style navigation with boundary checks
- ✅ Status bar feedback for operations

### Task 16: Context Menus - ✅ COMPLETE
**Goal**: Right-click operations
**Status**: ✅ Fully implemented
**Details**:
- ✅ 13 operations: Cut/Copy/Paste, Clear, Insert/Delete rows/columns
- ✅ MenuFlyout with emoji icons
- ✅ Clipboard integration

### Task 10: Theme System - ✅ COMPLETE
**Goal**: Visual customization
**Status**: ✅ Fully implemented
**Details**:
- ✅ Application themes (Light/Dark/System)
- ✅ Cell visual state themes (4 variants)
- ✅ Theme preview in settings

### Task 17: Settings Persistence - ✅ COMPLETE
**Goal**: Save user preferences
**Status**: ✅ Fully implemented
**Details**:
- ✅ Window size and position persistence
- ✅ Panel states (visibility, widths)
- ✅ Theme preferences
- ✅ Recent workbooks list (up to 10)
- ✅ JSON storage in %LocalAppData%\AiCalc

### Task 18: Undo/Redo System - ✅ COMPLETE
**Goal**: Command history management
**Status**: ✅ Fully implemented
**Details**:
- ✅ Stack-based command pattern (50-action limit)
- ✅ Tracks value, formula, format, mode changes
- ✅ Keyboard shortcuts: Ctrl+Z (Undo), Ctrl+Y (Redo)
- ✅ Status messages with feedback
- ✅ Automatic recording in cell changes

### Task 19: Formula Syntax Highlighting - ✅ COMPLETE
**Goal**: Visual formula feedback
**Status**: ✅ Fully implemented
**Details**:
- ✅ Real-time tokenization (Functions, Cell Refs, Strings, Numbers, Operators)
- ✅ Token counting display: "💡 2 functions, 3 cell refs"
- ✅ Sheet reference support (Sheet1!A1)
- ✅ Updates as user types

### Task 14B: Resizable Panels - ⏭️ SKIPPED
**Goal**: Flexible UI layout
**Status**: Skipped due to WinUI 3 XAML compiler bugs
**Dependencies**: None
**Details**:
- ⏭️ GridSplitter triggers XamlCompiler.exe errors
- ⏭️ Framework limitation, not implementation issue
- Workaround: Fixed panel sizes work well

---

### Task 15: Rich Cell Editing Dialog
**Goal**: Full-featured editor for complex cell types
**Dependencies**: Task 6
**Details**:
- Double-click Markdown cell → open Markdown editor dialog
- Split view: editor on left, preview on right
- Syntax highlighting for code blocks
- Support for Markdown tables, images, links
- JSON/XML editor with tree view and validation
- Table cell editor with grid view
- Image viewer with metadata display

**Questions**:
- Should the dialog be modal or can user edit multiple cells? modal
- Do you want a toolbar with formatting buttons for Markdown? yes

---

### Task 16: Right-Click Context Menu
**Goal**: Rich context-sensitive actions
**Dependencies**: Multiple earlier tasks
**Details**:
- Basic operations: Cut, Copy, Paste, Delete
- Insert/Delete Rows/Columns
- Format cell (background, text color, borders)
- Data Sources submenu (Task 18)
- Extract formula
- Clear contents vs Clear formatting
- Show cell history/audit trail

---

### Task 17: Pivot & Chart Cells
**Goal**: Data visualization components
**Dependencies**: Task 3 (table functions)
**Details**:
- Add `CellObjectType.Pivot` and `CellObjectType.Chart`
- Click pivot cell → open pivot configuration dialog
- Select source range, rows, columns, values, aggregation
- Click chart cell → open chart builder
- Support chart types: bar, line, pie, scatter
- Charts auto-update when source data changes
- Export charts as images

**Questions**:
- Which charting library do you prefer (OxyPlot, LiveCharts, etc.)?
please use your judgement, which is the most attractive, professional and flexible
- Should pivots support multiple aggregation functions?
yes

---

## Phase 6: Data Sources & External Connections

### Task 18: Data Sources Menu & Integration
**Goal**: Connect to external data sources
**Dependencies**: Task 11 (for connection management)
**Details**:
- Add "Data Sources" menu to main window
- Implement data source connections: Azure, AWS, GCP, SQL, REST APIs
- Create data access classes for each provider
- Right-click cell → "Load from Data Source"
- Support refresh operations (manual or scheduled)
- Add data source credentials management
- Implement query builder for databases

**Questions**:
- Which data sources are priority? (SQL Server, Azure Blob, S3, etc.) azure blob , azure sqldb
- Should queries be editable in a visual query builder? yes

---

### Task 19: Cell-Level Data Source Binding
**Goal**: Link individual cells to data sources
**Dependencies**: Task 18
**Details**:
- Right-click cell → "Bind to Data Source"
- Configure query/path for cell value
- Support refresh modes: manual, on open, scheduled
- Show binding indicator in cell
- Add refresh button in Cell Inspector
- Support parameter passing from other cells

---

## Phase 7: Python SDK & Scripting Integration

### Task 20: Python SDK - Local Environment Detection ✅ COMPLETE
**Goal**: Integrate with local Python runtime
**Dependencies**: None (can be parallel)
**Status**: ✅ Named Pipes IPC bridge, Python SDK client, environment detection, Settings UI all complete
**Details**:
- Detect installed Python environments (venv, conda, system) ✅ DONE
- Add Python environment selector in Settings ✅ DONE
- Implement Python SDK with functions: ✅ DONE
  - `connect()`: Connect to AiCalc from Python ✅ DONE
  - `get_value(cell_ref)`: Read cell value ✅ DONE
  - `set_value(cell_ref, value)`: Write cell value ✅ DONE
  - `run_function(name, args)`: Execute AiCalc function ✅ DONE
  - `get_range(range_ref)`: Read range as pandas DataFrame ⚠️ Partial (future enhancement)
  - `get_sheets()`: Get sheet list ✅ DONE
- Create Python package `aicalc-sdk` (installable via pip) ✅ DONE
- Settings UI with environment detection and SDK installation ✅ DONE

**Implementation Notes**:
- Using Named Pipes (\\.\pipe\AiCalc_Bridge) for secure IPC
- PythonBridgeService.cs: Server-side IPC handler with Byte mode transmission
- python-sdk/aicalc_sdk/client.py: Python client with pywin32
- JSON-based request/response protocol with case-insensitive deserialization
- Server starts automatically with WorkbookViewModel
- PythonEnvironmentDetector.cs: Registry, PATH, Conda, Venv detection (342 lines)
- Settings dialog Python tab: Environment selector, SDK status, test connection, install SDK
- UserPreferences.cs: PythonEnvironmentPath and PythonBridgeEnabled properties
- Verified working: set_value(), get_value(), get_sheets(), run_function()

**Questions**:
- Should the Python SDK use IPC, REST API, or COM interop?
✅ ANSWERED: Named Pipes (secure, no separate server, Windows-native)
- Do you want Jupyter Notebook integration? Not at this stage, if we have an SDK we can connect from Python within Jupyter I think.

---

### Task 21: Python Function Discovery & Execution ✅ COMPLETE
**Goal**: Load custom Python functions into AiCalc
**Dependencies**: Task 20
**Status**: ✅ Discovery, typed execution, hot reload, and VS Code editing integrated
**Details**:
- Scan Python environment for functions with `@aicalc_function` decorator
- Auto-register Python functions in FunctionRegistry
- Execute Python functions from cells
- Support parameter type hints and validation
- Return values automatically converted to appropriate CellObjectType
- Hot reload: detect Python file changes and update registry
- Add "Open in VS Code" button for Python script cells

**Questions**:
- Should Python functions run in-process or in separate process? I think a seperate process is more secure and reliable, however if we can do in-process that would be good. I dont want a seperate window appearing
- Do you want debugging support (attach debugger)? Yes, but if complicated we can work without

---

### Task 22: Cloud-Deployed Python Scripts
**Goal**: Deploy and execute parameterized Python scripts in cloud
**Dependencies**: Task 21, Task 18
**Details**:
- Support Azure Functions, AWS Lambda deployment
- Parameterize scripts from cell values
- Deploy script from AiCalc UI
- Monitor execution logs and status
- Handle async execution with callbacks
- Cache results for expensive operations

**Questions**:
- Which cloud providers are priority? Azure
- Should there be a script marketplace/sharing feature? Yes, for now I was thinking just pointing to a Github repo. We could do a proper marketplace in the future so future proofing would be good.

---

## Phase 8: Advanced Features & Polish

### Task 23: Enhanced Cell Inspector with Tabs
**Goal**: Organize Cell Inspector features
**Dependencies**: Multiple earlier tasks
**Details**:
- Add tabs: Value, Formula, Automation, Notes, Data Source, History
- Value tab: Rich editor based on cell type
- Formula tab: Formula builder with intellisense
- Automation tab: Triggers and scheduling
- Notes tab: Rich text notes with formatting
- Data Source tab: Connection configuration
- History tab: Show value changes over time with timestamps

---

### Task 24: Workbook Automation System
**Goal**: Advanced automation and triggers
**Dependencies**: Task 10, Task 9
**Details**:
- Expand automation modes beyond current Manual/OnEdit/OnLoad/Continuous
- Add triggers: Time-based, Data-change-based, Event-based
- Create automation workflow builder (visual)
- Support conditional evaluation (if X then evaluate Y)
- Add automation log viewer
- Support macro recording for repetitive tasks

---

### Task 25: File Format & Persistence ✅ 80% Complete
**Goal**: Optimize save/load operations
**Dependencies**: None
**Status**: Core features complete - AutoSave UI, CSV export/import with file pickers
**Details**:
- ✅ Optimize JSON serialization (currently implemented)
- ✅ Add autosave feature with recovery (timer-based, 1-60 min intervals, backup files)
- ✅ AutoSave settings UI (enable/disable toggle, interval slider in Settings dialog)
- ✅ Support export to CSV with file save picker (single sheet, proper escaping, UTF-8)
- ✅ Support import from CSV with file open picker (creates new sheet, robust parsing)
- ✅ User preferences persistence (AutoSaveEnabled, AutoSaveIntervalMinutes)
- ⏳ Add binary format option for large workbooks
- ⏳ Implement incremental save (only changed cells)
- ⏳ Support export to Excel (.xlsx) - optional
- ⏳ Support import from Excel - optional
- ⏳ Add version control (track changes, rollback)

**Implementation**:
- `Services/AutoSaveService.cs` - Automatic workbook saving with dirty flag tracking
- `Services/CsvService.cs` - CSV export and import functionality
- `Models/UserPreferences.cs` - AutoSaveEnabled and AutoSaveIntervalMinutes properties
- `SettingsDialog.xaml` - AutoSave section with toggle and slider
- `SettingsDialog.xaml.cs` - AutoSave event handlers, preference loading/saving
- `WorkbookViewModel.cs` - SetAutoSaveEnabled/SetAutoSaveInterval methods, CSV commands
- `MainWindow.xaml` - Export CSV / Import CSV buttons with tooltips
- `MainWindow.xaml.cs` - File pickers for CSV export/import
- `CellViewModel.cs` - MarkAsUpdated() marks workbook as dirty for autosave

**Questions**:
- Should binary format be default or opt-in? 
- Do you want cloud sync support (OneDrive, Dropbox)? Yes if easy to implement. My assumption is the underlying file access would do this

---

### Task 26: Row/Column Operations
**Goal**: Manage grid structure
**Dependencies**: Task 7
**Details**:
- Insert/Delete rows and columns
- Hide/Show rows and columns
- Resize row height and column width (drag handles)
- Freeze rows/columns (split panes)
- Auto-fit column width based on content
- Group/Outline rows (collapsible sections)
- Right-click row/column headers for operations

---

### Task 27: Selection & Range Operations
**Goal**: Multi-cell operations
**Dependencies**: Task 5
**Details**:
- Implement click-drag selection
- Shift+Click for range selection
- Ctrl+Click for multi-selection
- Show selection count and sum in status bar
- Fill down/right operations
- Format painter (copy formatting)
- Apply formula to range
- Find & Replace across selection/sheet/workbook

---

### Task 28: Themes & Customization
**Goal**: Visual customization
**Dependencies**: None
**Details**:
- Create theme system (Dark, Light, High Contrast)
- Allow custom color schemes
- Configurable font family and size
- Cell style presets (bold, currency, percentage, date)
- Conditional formatting rules
- Grid line visibility options
- Save theme preferences per user

---

### Task 29: Community Functions System
**Goal**: Plugin/extension architecture
**Dependencies**: Task 21
**Details**:
- Create plugin API for third-party functions
- Function package format (.aicalc-plugin)
- Plugin marketplace/repository
- Version management and updates
- Security sandboxing for untrusted code
- Rating and review system
- One-click install from marketplace

**Questions**:
- Should plugins be .NET assemblies or Python scripts or both? Both, primary for now is Python.
- Do you want a plugin signing/certification system? That would be great

---

### Task 30: Plans & Monetization System
**Goal**: Implement tiered service model
**Dependencies**: Task 11, Task 12
**Details**:
- Build scaffolding for future, we will only implement standard for now.
- Create service tiers: Standard, Plus, Pro
- Implement usage tracking and quotas
- Token limits per tier
- Model access restrictions (GPT-4 for Pro only, etc.)
- Volume discounts for Pro tier
- Endpoint/provider limits
- License key validation system
- Usage dashboard and billing integration

**Questions**:
- Is this for commercial use or personal project? Both
- Should there be a free tier with limitations? Yes this is standard.


---

## Phase 9: Testing, Documentation & Deployment

### Task 31: Unit Testing
**Goal**: Comprehensive test coverage
**Dependencies**: None (ongoing)
**Details**:
- Write unit tests for all Models
- Test FunctionRegistry and FunctionRunner
- Test CellAddress parsing and formatting
- Test dependency graph (DAG) implementation
- Test Python SDK integration
- Mock AI service calls for testing
- Aim for 80%+ code coverage

---

### Task 32: Integration Testing
**Goal**: End-to-end testing
**Dependencies**: Task 31
**Details**:
- Test complete workflows (create workbook → edit cells → save → load)
- Test AI function execution with real services
- Test multi-threading and concurrency
- Test large workbook performance
- Test data source integration
- UI automation tests with WinAppDriver

---

### Task 33: Documentation
**Goal**: Complete user and developer documentation
**Dependencies**: All feature tasks
**Details**:
- Update README.md with feature overview
- Create user guide with screenshots
- Document all built-in functions with examples
- Create Python SDK documentation
- API reference for plugin developers
- Video tutorials for key features
- FAQ and troubleshooting guide

---

### Task 34: Performance Optimization
**Goal**: Optimize for large workbooks
**Dependencies**: Task 9
**Details**:
- Implement cell virtualization (only render visible cells)
- Lazy evaluation (don't calculate hidden cells)
- Memory profiling and optimization
- Optimize JSON serialization/deserialization
- Database-backed cell storage for massive workbooks
- Benchmark and profile critical paths

---

### Task 35: Deployment & Distribution
**Goal**: Package and distribute application
**Dependencies**: All tasks
**Details**:
- Create MSIX installer for Windows
- Code signing certificate
- Auto-update mechanism
- Crash reporting and telemetry (opt-in)
- Setup wizard for first-time users
- Sample workbooks and templates
- Microsoft Store submission (optional)

---

## Implementation Notes

### Recommended Implementation Order:
1. **Quick Wins** (1-2 weeks): Tasks 1-3 (Core object system)
2. **Usability** (1-2 weeks): Tasks 4-6 (Navigation and editing)
3. **Core Engine** (2-3 weeks): Tasks 8-10 (DAG and multi-threading)
4. **AI Integration** (2-3 weeks): Tasks 11-13 (AI functions)
5. **Advanced Features** (3-4 weeks): Tasks 14-22 (UI, Python, Data sources)
6. **Polish** (2-3 weeks): Tasks 23-30 (Inspector, automation, themes)
7. **Quality & Ship** (2-3 weeks): Tasks 31-35 (Testing, docs, deployment)

### Priority Levels:
- **P0 (Critical)**: Tasks 1, 2, 4, 5, 8, 9, 11, 12
- **P1 (High)**: Tasks 3, 6, 7, 10, 13, 20, 21, 25, 26
- **P2 (Medium)**: Tasks 14, 15, 16, 17, 18, 19, 23, 27, 28, 31-35
- **P3 (Nice to have)**: Tasks 22, 24, 29, 30

### Key Decision Points:
Please review the questions marked throughout and provide guidance on:
1. Cell type implementation details
2. AI service provider priorities
3. Data source priorities
4. Python integration approach
5. Monetization/licensing plans (if applicable)

---

## Next Steps

1. Review this task list and provide feedback
2. Answer the questions noted in each task
3. Confirm priority order
4. We'll then proceed with Task 1 implementation

**Ready to start building?** 🚀
