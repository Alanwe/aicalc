# AiCalc Codebase Reference

**Last Updated**: October 18, 2025  
**Framework**: .NET 8.0 + Windows App SDK 1.4 (WinUI 3)  
**Architecture**: MVVM Pattern

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Solution Structure](#solution-structure)
3. [Core Architecture](#core-architecture)
4. [Models Layer](#models-layer)
5. [Services Layer](#services-layer)
6. [ViewModels Layer](#viewmodels-layer)
7. [UI Layer](#ui-layer)
8. [Key Dependencies](#key-dependencies)
9. [Data Flow](#data-flow)
10. [Extension Points](#extension-points)
11. [Build & Run](#build--run)

---

## Project Overview

AiCalc is an AI-native spreadsheet application built with WinUI 3. It extends traditional spreadsheet capabilities with:
- AI-powered cell functions (image captioning, text-to-image, etc.)
- Multi-threaded formula evaluation with dependency management
- Rich cell types (text, number, image, video, PDF, code, table, JSON, etc.)
- Configurable AI service connections (Azure OpenAI, Ollama, OpenAI, Anthropic)
- Python SDK integration for custom functions

**Primary Use Case**: Create intelligent spreadsheets where cells can contain any data type and leverage AI models for transformation, analysis, and generation.

---

## Solution Structure

```
AiCalc.sln                              # Main solution file
├── src/
│   └── AiCalc.WinUI/                   # Main WinUI 3 application
│       ├── AiCalc.WinUI.csproj         # Project file (net8.0-windows10.0.19041.0)
│       ├── App.xaml[.cs]               # Application entry point
│       ├── MainWindow.xaml[.cs]        # Primary UI shell
│       ├── Models/                     # Data models
│       ├── ViewModels/                 # MVVM view models
│       ├── Services/                   # Business logic & engines
│       ├── Converters/                 # XAML value converters
│       ├── Themes/                     # Visual themes
│       └── [Dialogs]                   # UI dialogs (Settings, Connections, etc.)
├── tests/
│   └── AiCalc.Tests/                   # xUnit test project (59 passing tests)
├── python-sdk/                         # Python SDK for custom functions (via named pipes)
├── docs/                               # Implementation notes & phase summaries
└── [Root Files]                        # README, tasks, status, guides
```

**Active Project**: `src/AiCalc.WinUI/` (all development happens here)

---

## Core Architecture

### MVVM Pattern

```
┌─────────────────────────────────────────────────────────────────┐
│                         UI Layer (XAML)                          │
│  MainWindow.xaml  │  Dialogs  │  Themes  │  Converters          │
└───────────────────────┬─────────────────────────────────────────┘
                        │ Binds to
                        ▼
┌─────────────────────────────────────────────────────────────────┐
│                      ViewModels Layer                            │
│  WorkbookViewModel  │  SheetViewModel  │  CellViewModel          │
│  - ObservableObject  │  - Commands     │  - INotifyPropertyChanged│
└───────────────────────┬─────────────────────────────────────────┘
                        │ Uses
                        ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Services Layer                              │
│  FunctionRunner  │  EvaluationEngine  │  DependencyGraph        │
│  FunctionRegistry  │  AIServiceRegistry  │  CredentialService   │
└───────────────────────┬─────────────────────────────────────────┘
                        │ Operates on
                        ▼
┌─────────────────────────────────────────────────────────────────┐
│                        Models Layer                              │
│  CellDefinition  │  SheetDefinition  │  WorkbookDefinition      │
│  CellValue  │  CellAddress  │  WorkspaceConnection              │
└─────────────────────────────────────────────────────────────────┘
```

**Key Principles**:
- UI never directly manipulates models
- ViewModels expose commands and properties for UI binding
- Services encapsulate business logic and evaluation
- Models are pure data structures (no logic)

---

## Models Layer

**Location**: `src/AiCalc.WinUI/Models/`

### Core Models

#### `CellAddress.cs` (178 lines)
**Purpose**: Parse and format cell references (A1, AA100, Sheet1!B2)

**Key Methods**:
```csharp
public static CellAddress Parse(string reference)      // "A1" → CellAddress
public string ToRelativeString()                       // CellAddress → "A1"
public string ToAbsoluteString()                       // CellAddress → "Sheet1!A1"
public static string ColumnIndexToName(int index)      // 0 → "A", 25 → "Z", 26 → "AA"
public static int ColumnNameToIndex(string name)       // "AA" → 26
```

**Properties**:
- `Sheet` (string?): Sheet name (null if current sheet)
- `Row` (int): Zero-based row index
- `Column` (int): Zero-based column index

**Dependencies**: None (standalone)

---

#### `CellValue.cs` (108 lines)
**Purpose**: Typed cell values with display formatting

**Structure**:
```csharp
public CellObjectType ObjectType { get; }     // Text, Number, Image, Table, etc.
public string RawValue { get; }               // Underlying value (file path, JSON, etc.)
public string DisplayValue { get; }           // User-facing formatted value
```

**Factory Methods**:
```csharp
public static CellValue Empty()
public static CellValue FromNumber(double value)
public static CellValue FromText(string text)
public static CellValue FromImage(string imagePath)
public static CellValue FromDirectory(string path)
public static CellValue FromTable(string[,] data)
public static CellValue FromError(string message)
```

**Dependencies**: `CellObjectType` enum

---

#### `CellDefinition.cs` (60 lines)
**Purpose**: Complete cell state (value + formula + metadata)

**Properties**:
```csharp
public CellAddress Address { get; }
public CellValue Value { get; set; }
public string? Formula { get; set; }                   // e.g., "=SUM(A1:A10)"
public CellAutomationMode AutomationMode { get; set; } // Manual, OnEdit, OnLoad, Continuous
public string? Notes { get; set; }
public CellFormat? Format { get; set; }                // Background, foreground, borders
```

**Dependencies**: `CellAddress`, `CellValue`, `CellAutomationMode`, `CellFormat`

---

#### `CellObjectType.cs` (31 types)
**Purpose**: Enum defining all supported cell types

**Categories**:
- **Primitives**: Empty, Text, Number, Markdown
- **Media**: Image, Video, Audio
- **File System**: Directory, File
- **Structured Data**: Table, Json, Xml
- **Documents**: Pdf, PdfPage
- **Code**: CodePython, CodeCSharp, CodeJavaScript, CodeTypeScript, CodeHtml, CodeCss
- **Other**: Script, Error

---

#### `SheetDefinition.cs` / `WorkbookDefinition.cs`
**Purpose**: Hierarchical structure for workbook data

**Structure**:
```csharp
WorkbookDefinition
  ├── Sheets: List<SheetDefinition>
  └── Settings: WorkbookSettings

SheetDefinition
  ├── Name: string
  ├── Cells: List<CellDefinition>
  ├── RowCount: int
  └── ColumnCount: int
```

---

#### `WorkspaceConnection.cs` (Phase 4)
**Purpose**: AI service connection configuration

**Key Properties**:
```csharp
public string Id { get; }
public string Name { get; set; }                // "My Azure OpenAI"
public string Provider { get; set; }            // "AzureOpenAI", "Ollama", "OpenAI"
public string Endpoint { get; set; }            // API endpoint URL
public string? ApiKeyEncrypted { get; set; }    // DPAPI-encrypted API key
public string? Model { get; set; }              // "gpt-4", "llama2"
public string? VisionModel { get; set; }        // "gpt-4-vision", "llava"
public string? ImageModel { get; set; }         // "dall-e-3"
public int TimeoutSeconds { get; set; } = 100;
public double Temperature { get; set; } = 0.7;
public bool IsActive { get; set; }
```

**Methods**:
```csharp
public WorkspaceConnection Clone()             // Safe copy for editing
```

**Dependencies**: `CredentialService` for encryption/decryption

---

#### `CellVisualState.cs` (Phase 3)
**Purpose**: Enum for cell UI states (border colors)

**Values**:
- `Normal`: Default state
- `JustUpdated`: Green flash (2 seconds after update)
- `Calculating`: Orange spinner during evaluation
- `Stale`: Blue border (needs recalculation)
- `ManualUpdate`: Orange indicator (manual mode)
- `Error`: Red border (formula error)
- `InDependencyChain`: Gold highlight (when dependent cell selected)

---

#### `FunctionSuggestion.cs` (Phase 4)
**Purpose**: Intellisense/autocomplete metadata for functions

**Properties**:
```csharp
public string Name { get; set; }               // "SUM"
public string Description { get; set; }        // "Adds numbers"
public string Signature { get; set; }          // "SUM(values)"
public FunctionCategory Category { get; set; } // Math, AI, Text, etc.
public string CategoryGlyph { get; set; }      // "➕", "🤖", etc.
public string TypeHint { get; set; }           // "Input: Image / Video"
public string ProviderHint { get; set; }       // "Azure OpenAI · gpt-4-vision"
public bool HasTypeHint { get; }
public bool HasProviderHint { get; }
```

---

### Model Dependencies Graph

```
WorkbookDefinition
  └─> SheetDefinition[]
       └─> CellDefinition[]
            ├─> CellAddress
            ├─> CellValue
            │    └─> CellObjectType (enum)
            ├─> CellAutomationMode (enum)
            └─> CellFormat

WorkbookSettings
  └─> WorkspaceConnection[]
       └─> (encrypted API keys)

CellViewModel (ViewModel)
  └─> CellVisualState (enum) - for UI rendering
```

---

## Services Layer

**Location**: `src/AiCalc.WinUI/Services/`

### Core Services

#### `FunctionRegistry.cs` (683 lines)
**Purpose**: Central registry of all available functions (built-in, AI, community)

**Key Methods**:
```csharp
public void Register(FunctionDescriptor descriptor)
public bool TryGet(string name, out FunctionDescriptor descriptor)
public IEnumerable<FunctionDescriptor> GetFunctionsForTypes(params CellObjectType[] types)
public IEnumerable<FunctionDescriptor> GetFunctionsByCategory(FunctionCategory category)
```

**Built-in Functions** (25+ functions):
- **Math**: SUM, AVERAGE, COUNT, MIN, MAX, ROUND
- **Text**: CONCAT, UPPER, LOWER, TRIM, LEN, SPLIT, REPLACE
- **DateTime**: NOW, TODAY, DATE, TIME, DATEADD, DATEDIFF
- **File**: FILE_SIZE, FILE_EXTENSION, FILE_READ
- **Directory**: DIRECTORY_TO_TABLE, DIR_SIZE
- **Table**: TABLE_FILTER, TABLE_SORT, TABLE_AGGREGATE
- **Image**: (helpers for IMAGE_TO_CAPTION)
- **PDF**: PDF_PAGE_COUNT, PDF_PAGE_TEXT
- **AI**: IMAGE_TO_CAPTION, TEXT_TO_IMAGE, ANALYZE_IMAGE, SUMMARIZE_TEXT, EXTRACT_ENTITIES, TRANSLATE, CODE_EXPLAIN, CODE_GENERATE, CODE_FIX

**Function Registration Example**:
```csharp
Register(new FunctionDescriptor(
    "SUM",
    "Adds a series of numbers.",
    async ctx => {
        var sum = ctx.Arguments.Sum(cell => double.Parse(cell.DisplayValue));
        return new FunctionExecutionResult(CellValue.FromNumber(sum));
    },
    FunctionCategory.Math,
    new FunctionParameter("values", "Range of values", CellObjectType.Number)
));
```

**Dependencies**: `FunctionDescriptor`, `FunctionParameter`, `CellObjectType`, `WorkbookViewModel`, `AIServiceRegistry`

---

#### `FunctionRunner.cs` (322 lines)
**Purpose**: Parses and executes cell formulas

**Key Methods**:
```csharp
public async Task<CellValue> EvaluateAsync(
    string formula, 
    CellViewModel cell, 
    SheetViewModel sheet, 
    WorkbookViewModel workbook)
```

**Formula Parsing**:
- Regex pattern: `@"^=([A-Z_]+)\((.*)\)$"`
- Extracts function name and arguments
- Resolves cell references (A1, B2:B10, Sheet2!C3)
- Handles ranges (expands A1:A10 into individual cells)
- Detects AI functions and routes to AI services

**AI Function Execution** (Phase 4):
```csharp
private async Task<FunctionExecutionResult> ExecuteAIFunctionAsync(
    string functionName,
    List<CellViewModel> args,
    WorkbookViewModel workbook)
{
    var connection = workbook.Settings.Connections.FirstOrDefault(c => c.IsActive);
    var client = App.AIServices.CreateClient(connection);
    var result = await client.ExecuteAsync(functionName, args);
    return new FunctionExecutionResult(result, diagnostics: aiResponse.Metadata);
}
```

**Dependencies**: `FunctionRegistry`, `DependencyGraph`, `CellViewModel`, `AIServiceRegistry`

---

#### `DependencyGraph.cs` (Phase 3 - 238 lines)
**Purpose**: Tracks cell dependencies for efficient evaluation (DAG)

**Key Methods**:
```csharp
public void AddDependency(CellAddress dependentCell, CellAddress referencedCell)
public IReadOnlyList<CellAddress> GetDependents(CellAddress cell)
public IReadOnlyList<CellAddress> GetDependencies(CellAddress cell)
public bool HasCircularReference(CellAddress startCell, out List<CellAddress> cycle)
public List<CellAddress> TopologicalSort()
public void RemoveCell(CellAddress cell)
public void Clear()
```

**Data Structure**:
```csharp
private class DependencyNode
{
    public CellAddress Address { get; }
    public HashSet<CellAddress> Dependencies { get; }    // Cells this cell references
    public HashSet<CellAddress> Dependents { get; }      // Cells that reference this cell
}

private Dictionary<CellAddress, DependencyNode> _nodes;
```

**Circular Reference Detection**:
- Depth-first search with visited tracking
- Returns cycle path if found
- Example: A1 → B1 → C1 → A1 (cycle detected)

**Topological Sort**:
- Returns evaluation order (dependencies first)
- Used by `EvaluationEngine` for parallel execution

**Dependencies**: `CellAddress`

---

#### `EvaluationEngine.cs` (Phase 3 - 314 lines)
**Purpose**: Multi-threaded cell evaluation with cancellation support

**Key Methods**:
```csharp
public async Task<bool> EvaluateCellAsync(CellViewModel cell, CancellationToken ct)

public async Task<EvaluationResult> EvaluateAllAsync(
    Dictionary<CellAddress, CellViewModel> cells,
    CancellationToken ct,
    IProgress<int>? progress = null)

public async Task<EvaluationResult> EvaluateDependentsAsync(
    CellAddress changedCell,
    Dictionary<CellAddress, CellViewModel> cells,
    CancellationToken ct)
```

**Configuration**:
```csharp
public int MaxDegreeOfParallelism { get; set; }     // Default: Environment.ProcessorCount
public int DefaultTimeoutSeconds { get; set; }       // Default: 100
```

**Evaluation Flow**:
1. Get topological sort from `DependencyGraph`
2. Filter by automation mode (skip `Manual`)
3. Group cells by dependency level (independent cells in parallel)
4. Execute using `Parallel.ForEachAsync` with cancellation support
5. Report progress via `IProgress<int>`

**Result Structure**:
```csharp
public record EvaluationResult(
    int TotalCells,
    int SuccessCount,
    int ErrorCount,
    int SkippedCount,
    TimeSpan Duration,
    List<(CellAddress Address, string Error)> Errors);
```

**Dependencies**: `DependencyGraph`, `FunctionRunner`, `CellViewModel`

---

#### `FunctionDescriptor.cs` (140 lines)
**Purpose**: Metadata for function definition

**Structure**:
```csharp
public class FunctionDescriptor
{
    public string Name { get; }
    public string Description { get; }
    public FunctionCategory Category { get; }
    public IReadOnlyList<FunctionParameter> Parameters { get; }
    public CellObjectType[] ApplicableTypes { get; }    // Cell types this function accepts
    public Func<FunctionEvaluationContext, Task<FunctionExecutionResult>> Handler { get; }
    
    public bool CanAccept(params CellObjectType[] types) // Type checking
}

public class FunctionParameter
{
    public string Name { get; }
    public string Description { get; }
    public CellObjectType ExpectedType { get; }
    public CellObjectType[] AcceptableTypes { get; }    // Supports polymorphism
    public bool IsOptional { get; }
    
    public bool CanAccept(CellObjectType type)
}

public enum FunctionCategory
{
    Math, Text, DateTime, File, Directory, Table, 
    Image, Video, Pdf, Data, AI, Contrib
}

public record FunctionExecutionResult(
    CellValue Value,
    string? Diagnostics = null,
    CellValue[,]? SpillRange = null,                    // For array formulas
    IReadOnlyList<CellAddress>? ReferencedCells = null,
    AIResponse? AiResponse = null);                     // Phase 4: AI metadata

public record FunctionEvaluationContext(
    WorkbookViewModel Workbook,
    SheetViewModel Sheet,
    IReadOnlyList<CellViewModel> Arguments,
    string RawFormula);
```

**Dependencies**: `CellValue`, `CellObjectType`, `CellAddress`, `WorkbookViewModel`, `AIResponse`

---

### AI Services (Phase 4)

**Location**: `src/AiCalc.WinUI/Services/AI/`

#### `AIServiceRegistry.cs` (134 lines)
**Purpose**: Manage AI client instances and connections

**Key Methods**:
```csharp
public void RegisterConnection(WorkspaceConnection connection)
public void UnregisterConnection(string connectionId)
public WorkspaceConnection? GetDefaultConnection()
public IAIServiceClient CreateClient(WorkspaceConnection? connection)
public async Task<bool> TestConnectionAsync(WorkspaceConnection connection)
```

**Supported Providers**:
- `AzureOpenAIClient`: Azure OpenAI Service
- `OllamaClient`: Local Ollama models
- *(Future: OpenAI, Anthropic, Google, etc.)*

**Dependencies**: `WorkspaceConnection`, `IAIServiceClient`, `CredentialService`

---

#### `IAIServiceClient.cs` (Interface)
**Purpose**: Common interface for all AI providers

**Methods**:
```csharp
Task<AIResponse> GenerateTextAsync(string prompt, CancellationToken ct)
Task<AIResponse> GenerateImageAsync(string prompt, CancellationToken ct)
Task<AIResponse> AnalyzeImageAsync(byte[] imageData, string prompt, CancellationToken ct)
Task<bool> TestConnectionAsync()
```

**AIResponse Structure**:
```csharp
public record AIResponse(
    string Content,
    int TokensUsed,
    TimeSpan Duration,
    string? Model = null,
    Dictionary<string, object>? Metadata = null);
```

---

#### `AzureOpenAIClient.cs` (245 lines)
**Purpose**: Azure OpenAI API integration

**Features**:
- Text generation (gpt-4, gpt-3.5-turbo)
- Vision analysis (gpt-4-vision)
- Image generation (dall-e-3)
- Configurable timeout, temperature, max tokens
- Retry logic (3 attempts default)

**Dependencies**: `System.Net.Http`, `System.Text.Json`, `WorkspaceConnection`

---

#### `OllamaClient.cs` (198 lines)
**Purpose**: Local Ollama model integration

**Features**:
- Text generation (llama2, mistral, etc.)
- Vision analysis (llava)
- Streaming support (future)
- Local inference (no API key required)

**Endpoint**: `http://localhost:11434/api/generate`

---

#### `CredentialService.cs` (40 lines)
**Purpose**: Secure API key storage using Windows DPAPI

**Key Methods**:
```csharp
public static string Encrypt(string plaintext)         // CurrentUser scope
public static string Decrypt(string encryptedBase64)
```

**Security**:
- Uses `ProtectedData.Protect` (Windows DPAPI)
- Keys encrypted per Windows user account
- Base64 encoding for storage in JSON

**Dependencies**: `System.Security.Cryptography`

---

### Service Dependencies Graph

```
EvaluationEngine
  ├─> DependencyGraph (topological sort)
  ├─> FunctionRunner (execute formulas)
  └─> CellViewModel (update state)

FunctionRunner
  ├─> FunctionRegistry (lookup functions)
  ├─> AIServiceRegistry (route AI calls)
  └─> DependencyGraph (track references)

FunctionRegistry
  └─> FunctionDescriptor[] (function definitions)

AIServiceRegistry
  ├─> WorkspaceConnection[] (connection configs)
  ├─> CredentialService (decrypt API keys)
  └─> IAIServiceClient implementations
       ├─> AzureOpenAIClient
       └─> OllamaClient

DependencyGraph
  └─> CellAddress (node keys)
```

---

## ViewModels Layer

**Location**: `src/AiCalc.WinUI/ViewModels/`

All ViewModels inherit from `BaseViewModel` (CommunityToolkit.Mvvm.ComponentModel):
- `ObservableObject` for `INotifyPropertyChanged`
- `[ObservableProperty]` source generator for property backing fields
- `[RelayCommand]` source generator for commands

---

#### `WorkbookViewModel.cs` (504 lines)
**Purpose**: Top-level workbook management and orchestration

**Key Properties**:
```csharp
public ObservableCollection<SheetViewModel> Sheets { get; }
public SheetViewModel? SelectedSheet { get; set; }
public WorkbookSettings Settings { get; set; }
public string Title { get; set; }
public CellViewModel? ActiveCell { get; set; }

// Core Services
public FunctionRegistry FunctionRegistry { get; }
public FunctionRunner FunctionRunner { get; }
public DependencyGraph DependencyGraph { get; }
public EvaluationEngine EvaluationEngine { get; }
```

**Key Commands**:
```csharp
[RelayCommand] public async Task AddSheet()
[RelayCommand] public async Task RemoveSheet(SheetViewModel sheet)
[RelayCommand] public async Task SaveAsync(string path)
[RelayCommand] public async Task LoadAsync(string path)
[RelayCommand] public async Task RecalculateAll()
```

**Settings Synchronization** (Phase 4):
```csharp
private void AttachSettings()
{
    Settings.Connections.CollectionChanged += OnConnectionsCollectionChanged;
}

private void OnConnectionsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
{
    SynchronizeConnectionsToRegistry();  // Sync to AIServiceRegistry
}

private void SynchronizeConnectionsToRegistry()
{
    // Ensure all connections have encrypted API keys
    // Register with AIServiceRegistry
}
```

**Dependencies**: All core services, `SheetViewModel`, `WorkbookSettings`

---

#### `SheetViewModel.cs` (287 lines)
**Purpose**: Individual sheet operations and cell management

**Key Properties**:
```csharp
public string Name { get; set; }
public int RowCount { get; set; } = 50
public int ColumnCount { get; set; } = 26
public Dictionary<CellAddress, CellViewModel> Cells { get; }
public ObservableCollection<RowViewModel> Rows { get; }
public List<string> ColumnHeaders { get; }        // A, B, C, ..., AA, AB, ...
public WorkbookViewModel Workbook { get; }
```

**Key Methods**:
```csharp
public CellViewModel GetOrCreateCell(int row, int col)
public CellViewModel? GetCell(int row, int col)
public CellViewModel? GetCell(CellAddress address)
public void RemoveCell(CellAddress address)
public async Task EvaluateAllCellsAsync()
public async Task EvaluateCellAsync(CellViewModel cell)
```

**Cell Management**:
- Lazy cell creation (cells created on first access)
- Sparse storage (Dictionary, not 2D array)
- Automatic row/column expansion

**Dependencies**: `WorkbookViewModel`, `RowViewModel`, `CellViewModel`, `EvaluationEngine`

---

#### `CellViewModel.cs` (Phase 3/4 - 425 lines)
**Purpose**: Individual cell state and operations

**Key Properties**:
```csharp
public CellAddress Address { get; }
public CellValue Value { get; set; }
public string DisplayValue { get; }               // Formatted for UI
public string? Formula { get; set; }
public CellAutomationMode AutomationMode { get; set; }
public CellVisualState VisualState { get; set; }  // Phase 3: Border colors
public string? Notes { get; set; }
public CellFormat? Format { get; set; }
public bool IsSelected { get; set; }
public bool IsEditing { get; set; }
public bool IsEvaluating { get; set; }
public ObservableCollection<CellHistoryEntry> History { get; }  // Audit trail
```

**Key Methods**:
```csharp
public async Task EvaluateAsync(CancellationToken ct)
public void MarkAsStale()                         // Phase 3: Visual feedback
public void MarkAsCalculating()
public void MarkAsUpdated()
public void MarkAsError()
public void AppendHistory(string reason, string? details = null)  // Phase 4: AI diagnostics
```

**Visual State Management** (Phase 3):
```csharp
public void MarkAsUpdated()
{
    VisualState = CellVisualState.JustUpdated;
    Task.Delay(2000).ContinueWith(_ => {
        if (VisualState == CellVisualState.JustUpdated)
            VisualState = CellVisualState.Normal;
    });
}
```

**History Tracking** (Phase 4):
```csharp
public void AppendHistory(string reason, string? details = null)
{
    History.Add(new CellHistoryEntry(
        DateTime.Now,
        Value.DisplayValue,
        reason,
        details  // AI diagnostics: "Model: gpt-4-vision | Tokens: 150 | Latency: 1.2s"
    ));
}
```

**Dependencies**: `SheetViewModel`, `CellValue`, `CellAddress`, `CellVisualState`, `EvaluationEngine`

---

#### `RowViewModel.cs` (45 lines)
**Purpose**: Row-level operations (row height pending; hide/unhide supported)

**Properties**:
```csharp
public int RowIndex { get; }
public SheetViewModel Sheet { get; }
public ObservableCollection<CellViewModel> Cells { get; }
```

---

### ViewModel Dependencies Graph

```
WorkbookViewModel
  ├─> SheetViewModel[]
  │    ├─> RowViewModel[]
  │    │    └─> CellViewModel[]
  │    └─> Dictionary<CellAddress, CellViewModel>
  ├─> WorkbookSettings
  │    └─> WorkspaceConnection[]
  ├─> FunctionRegistry
  ├─> FunctionRunner
  ├─> DependencyGraph
  └─> EvaluationEngine

CellViewModel
  └─> CellHistoryEntry[] (ObservableCollection)
```

---

## UI Layer

**Location**: `src/AiCalc.WinUI/`

### Main Application

#### `App.xaml[.cs]` (115 lines)
**Purpose**: Application entry point and global services

**Key Responsibilities**:
- Initialize `AIServiceRegistry` (global singleton)
- Load themes from `Themes/CellStateThemes.xaml`
- Create main window
- Handle unhandled exceptions

**Global Services**:
```csharp
public static AIServiceRegistry AIServices { get; private set; }
```

**Startup Flow**:
```csharp
protected override void OnLaunched(LaunchActivatedEventArgs args)
{
    AIServices = new AIServiceRegistry();
    
    m_window = new Window();
    m_window.Content = new MainWindow();
    m_window.Activate();
}
```

---

#### `MainWindow.xaml[.cs]` (Phase 4/5 - 2037 lines)
**Purpose**: Primary UI shell with spreadsheet, functions panel, inspector

**Structure**:
```
MainWindow (Page)
  ├─ Header (Title, Buttons: New Sheet, Recalculate, Save, Load, Settings)
  ├─ Main Layout (3-column grid)
  │   ├─ Functions Panel (left, collapsible)
  │   │   └─ Function list (categorized, searchable)
  │   ├─ Spreadsheet Grid (center)
  │   │   ├─ Sheet Tabs (TabView)
  │   │   └─ Cell Grid (ScrollViewer + dynamic Grid)
  │   └─ Cell Inspector (right, collapsible)
  │       ├─ Value editor
  │       ├─ Formula editor (with autocomplete)
  │       ├─ Automation mode selector
  │       ├─ Evaluate button
  │       └─ Notes editor
  └─ Status Bar (bottom)
```

**Key Features** (Phase 4):
- **Formula Autocomplete**: Popup with function suggestions, type hints, provider hints
- **Parameter Hints**: Tooltip showing function signature while typing
- **Contextual Suggestions**: Filters AI functions by cell type and available connections

**Key Methods**:
```csharp
private void ShowFormulaIntellisense(string text, int caretIndex)
{
    var searchText = ExtractFunctionNamePrefix(text, caretIndex);
    var suggestions = GetContextualSuggestions(searchText);
    FunctionAutocompleteList.ItemsSource = suggestions;
    FunctionAutocompletePopup.IsOpen = true;
}

private IEnumerable<FunctionSuggestion> GetContextualSuggestions(string searchText)
{
    var connection = App.AIServices.GetDefaultConnection();
    var currentCellType = _selectedCell?.Value.ObjectType ?? CellObjectType.Text;
    
    foreach (var descriptor in ViewModel.FunctionRegistry.Functions)
    {
        if (descriptor.Category == FunctionCategory.AI)
        {
            if (!IsFunctionSupportedByConnection(descriptor, connection))
                continue;
            if (!IsFunctionRelevantForCellType(descriptor, currentCellType))
                continue;
        }
        yield return CreateSuggestion(descriptor, connection);
    }
}
```

**Data Context**:
```csharp
public WorkbookViewModel ViewModel { get; private set; }
```

**Dependencies**: `WorkbookViewModel`, `FunctionSuggestion`, `AIServiceRegistry`

---

### Dialogs

#### `ServiceConnectionDialog.xaml[.cs]` (Phase 4 - 387 lines)
**Purpose**: Configure AI service connections

**Features**:
- Provider selection (Azure OpenAI, Ollama, OpenAI)
- Multi-model configuration (Text, Vision, Image)
- API key entry (PasswordBox → encrypted storage)
- Connection testing (real-time feedback)
- Performance settings (timeout, retries, temperature)
- Preset buttons (Azure, Ollama quick setup)

**Key Methods**:
```csharp
private async void TestConnection_Click(object sender, RoutedEventArgs e)
{
    var testConnection = _editConnection.Clone();  // Safe copy
    testConnection.ApiKeyEncrypted = CredentialService.Encrypt(_apiKeyPlainText);
    
    var success = await App.AIServices.TestConnectionAsync(testConnection);
    ConnectionStatusText.Text = success ? "✅ Connected" : "❌ Failed";
}
```

**Security**:
- API keys stored in `PasswordBox` (not logged)
- Encrypted before saving to model
- Clone used for test connections (original untouched until Save)

---

#### `SettingsDialog.xaml[.cs]` (226 lines)
**Purpose**: Application settings

**Sections**:
- **AI Connections**: Manage `WorkspaceConnection[]`
  - Add, Edit, Delete, Set Default
- **Evaluation**: (Future) Thread count, timeout configuration
- **Theme**: (Future) Light/Dark/High Contrast

---

#### `EvaluationSettingsDialog.xaml[.cs]` (97 lines)
**Purpose**: Per-cell evaluation settings

**Settings**:
- Automation mode (Manual, OnEdit, OnLoad, Continuous)
- Timeout override (per-cell)
- Retry count

---

### Converters

**Location**: `src/AiCalc.WinUI/Converters/`

#### `BooleanToVisibilityConverter.cs`
**Purpose**: Show/hide UI elements based on boolean

**Usage**:
```xaml
Visibility="{Binding IsSelected, Converter={StaticResource BoolToVisibilityConverter}}"
```

---

#### `CellVisualStateToBrushConverter.cs` (Phase 3)
**Purpose**: Map `CellVisualState` enum to border colors

**Mapping**:
- `Normal` → Transparent
- `JustUpdated` → LimeGreen
- `Calculating` → Orange
- `Stale` → DodgerBlue
- `ManualUpdate` → Orange
- `Error` → Crimson
- `InDependencyChain` → Gold

**Usage**:
```xaml
BorderBrush="{Binding VisualState, Converter={StaticResource VisualStateToBrushConverter}}"
```

---

#### `AutomationModeToGlyphConverter.cs`
**Purpose**: Display icons for automation modes

**Mapping**:
- `Manual` → "✋"
- `OnEdit` → "✏️"
- `OnLoad` → "📂"
- `Continuous` → "🔄"

---

### Themes

**Location**: `src/AiCalc.WinUI/Themes/`

#### `CellStateThemes.xaml` (Phase 3)
**Purpose**: Define cell visual state colors

**Resource Dictionaries**:
- Light Theme
- Dark Theme
- High Contrast Theme

**Color Resources**:
```xaml
<SolidColorBrush x:Key="CellStateJustUpdatedBrush" Color="#32CD32"/>
<SolidColorBrush x:Key="CellStateCalculatingBrush" Color="#FFA500"/>
<SolidColorBrush x:Key="CellStateStaleBrush" Color="#1E90FF"/>
<!-- etc. -->
```

**Usage**:
- Merged in `App.xaml` resource dictionaries
- Referenced by converters at runtime
- Theme switching via Settings dialog (future)

---

## Key Dependencies

### NuGet Packages

**Core Framework**:
```xml
<PackageReference Include="Microsoft.WindowsAppSDK" Version="1.4.231219000" />
<PackageReference Include="Microsoft.Windows.SDK.BuildTools" Version="10.0.22621.756" />
```

**MVVM Toolkit**:
```xml
<PackageReference Include="CommunityToolkit.Mvvm" Version="8.2.1" />
```

**Cryptography**:
```xml
<PackageReference Include="System.Security.Cryptography.ProtectedData" Version="8.0.0" />
```

**JSON**:
```xml
<PackageReference Include="System.Text.Json" Version="8.0.5" />
```

**Testing**:
```xml
<PackageReference Include="xunit" Version="2.5.0" />
<PackageReference Include="xunit.runner.visualstudio" Version="2.5.0" />
```

---

### .NET Namespaces

**UI**:
- `Microsoft.UI.Xaml` - WinUI 3 controls
- `Microsoft.UI.Xaml.Controls` - Buttons, TextBoxes, etc.
- `Microsoft.UI.Xaml.Data` - Value converters
- `Microsoft.UI.Xaml.Media` - Brushes, colors
- `Windows.UI` - Additional UI types

**MVVM**:
- `CommunityToolkit.Mvvm.ComponentModel` - ObservableObject, attributes
- `CommunityToolkit.Mvvm.Input` - RelayCommand

**Core .NET**:
- `System.Collections.ObjectModel` - ObservableCollection
- `System.Text.Json` - Serialization
- `System.Text.RegularExpressions` - Formula parsing
- `System.Threading.Tasks` - Async/await
- `System.Security.Cryptography` - DPAPI encryption
- `System.IO.Pipes` - Named pipes (Python SDK)

---

## Data Flow

### Formula Evaluation Flow

```
1. User enters formula: "=SUM(A1:A10)" in Cell Inspector
   └─> CellViewModel.Formula updated

2. User clicks "Evaluate Cell" or AutomationMode triggers
   └─> CellViewModel.EvaluateAsync() called

3. CellViewModel calls EvaluationEngine
   └─> EvaluationEngine.EvaluateCellAsync(cell)

4. EvaluationEngine calls FunctionRunner
   └─> FunctionRunner.EvaluateAsync(formula, cell, sheet, workbook)

5. FunctionRunner parses formula
   ├─> Extract function name: "SUM"
   ├─> Extract arguments: "A1:A10"
   └─> Resolve references: [CellViewModel(A1), CellViewModel(A2), ...]

6. FunctionRunner looks up function in FunctionRegistry
   └─> FunctionRegistry.TryGet("SUM", out descriptor)

7. FunctionRunner executes function handler
   └─> descriptor.Handler(context) returns FunctionExecutionResult

8. Result propagated back
   └─> CellViewModel.Value = result.Value
   └─> CellViewModel.MarkAsUpdated() (green flash)

9. DependencyGraph notified
   └─> Dependent cells marked as Stale (blue border)
```

---

### AI Function Flow (Phase 4)

```
1. User enters AI formula: "=IMAGE_TO_CAPTION(A1)" where A1 = image path
   └─> CellViewModel.Formula updated

2. FunctionRunner detects AI function
   └─> Checks if function name in AIFunctionNames list

3. FunctionRunner routes to AI service
   └─> ExecuteAIFunctionAsync(functionName, args, workbook)

4. Get active connection from WorkbookSettings
   └─> connection = workbook.Settings.Connections.FirstOrDefault(c => c.IsActive)

5. Create AI client via AIServiceRegistry
   └─> client = App.AIServices.CreateClient(connection)

6. Decrypt API key (if needed)
   └─> CredentialService.Decrypt(connection.ApiKeyEncrypted)

7. Load image data from file path (A1.Value.RawValue)
   └─> byte[] imageData = File.ReadAllBytes(imagePath)

8. Call AI service
   └─> AIResponse response = await client.AnalyzeImageAsync(imageData, "Caption this image")

9. Return result with diagnostics
   └─> FunctionExecutionResult(
       Value: CellValue.FromText(response.Content),
       Diagnostics: $"Model: {response.Model} | Tokens: {response.TokensUsed}",
       AiResponse: response)

10. Update cell and history
    └─> CellViewModel.Value = result.Value
    └─> CellViewModel.AppendHistory("AI Evaluation", result.Diagnostics)
```

---

### Save/Load Flow

```
Save:
  WorkbookViewModel.SaveAsync(path)
    └─> WorkbookDefinition definition = ToDefinition()
         ├─> Convert SheetViewModel[] to SheetDefinition[]
         ├─> Convert CellViewModel[] to CellDefinition[]
         └─> Include WorkbookSettings (with encrypted API keys)
    └─> string json = JsonSerializer.Serialize(definition)
    └─> File.WriteAllText(path, json)

Load:
  WorkbookViewModel.LoadAsync(path)
    └─> string json = File.ReadAllText(path)
    └─> WorkbookDefinition definition = JsonSerializer.Deserialize(json)
    └─> Sheets.Clear()
    └─> foreach (SheetDefinition sheetDef in definition.Sheets)
         ├─> SheetViewModel sheet = new SheetViewModel(this, sheetDef.Name)
         └─> foreach (CellDefinition cellDef in sheetDef.Cells)
              ├─> CellViewModel cell = sheet.GetOrCreateCell(cellDef.Address)
              └─> cell.Value = cellDef.Value
              └─> cell.Formula = cellDef.Formula
    └─> Settings = definition.Settings
    └─> SynchronizeConnectionsToRegistry()  // Register AI connections
```

---

## Extension Points

### 1. Adding New Functions

**Location**: `src/AiCalc.WinUI/Services/FunctionRegistry.cs`

**Steps**:
1. Define function descriptor in `RegisterBuiltIns()` or category method
2. Specify category, parameters, and applicable cell types
3. Implement handler (async Task<FunctionExecutionResult>)
4. Function automatically appears in UI Functions panel

**Example**:
```csharp
private void RegisterMathFunctions()
{
    Register(new FunctionDescriptor(
        "MULTIPLY",
        "Multiplies two numbers.",
        async ctx => {
            var a = double.Parse(ctx.Arguments[0].DisplayValue);
            var b = double.Parse(ctx.Arguments[1].DisplayValue);
            return new FunctionExecutionResult(CellValue.FromNumber(a * b));
        },
        FunctionCategory.Math,
        new FunctionParameter("a", "First number", CellObjectType.Number),
        new FunctionParameter("b", "Second number", CellObjectType.Number)
    ));
}
```

---

### 2. Adding New AI Providers

**Location**: `src/AiCalc.WinUI/Services/AI/`

**Steps**:
1. Create new class implementing `IAIServiceClient`
2. Implement methods: `GenerateTextAsync`, `GenerateImageAsync`, `AnalyzeImageAsync`, `TestConnectionAsync`
3. Register provider in `AIServiceRegistry.CreateClient()`:
   ```csharp
   return connection.Provider switch {
       "AzureOpenAI" => new AzureOpenAIClient(connection),
       "Ollama" => new OllamaClient(connection),
       "NewProvider" => new NewProviderClient(connection),  // Add here
       _ => throw new NotSupportedException()
   };
   ```
4. Add provider to `ServiceConnectionDialog` ComboBox

---

### 3. Adding New Cell Types

**Location**: `src/AiCalc.WinUI/Models/CellObjectType.cs`

**Steps**:
1. Add enum value to `CellObjectType`
2. Update `CellValue` factory methods if needed
3. Register type-specific functions in `FunctionRegistry`
4. Add UI handling in `MainWindow.xaml.cs` (cell rendering)
5. Update converters if special formatting needed

**Example**:
```csharp
// 1. Add enum
public enum CellObjectType {
    // ... existing types ...
    GeoLocation,  // New type for lat/long
}

// 2. Add factory method
public static CellValue FromGeoLocation(double lat, double lon)
{
    return new CellValue(
        CellObjectType.GeoLocation,
        $"{lat},{lon}",
        $"📍 {lat:F6}, {lon:F6}");
}

// 3. Register functions
Register(new FunctionDescriptor(
    "GEO_DISTANCE",
    "Calculate distance between two locations.",
    async ctx => { /* implementation */ },
    FunctionCategory.Data,
    new FunctionParameter("location1", "First location", CellObjectType.GeoLocation),
    new FunctionParameter("location2", "Second location", CellObjectType.GeoLocation)
));
```

---

### 4. Python SDK Integration

**Location**: `python-sdk/` and `src/AiCalc.WinUI/Services/PipeServer.cs`

**Architecture**:
- AiCalc hosts a named pipe server (`PipeServer.cs`)
- Python SDK connects via named pipe
- Bidirectional communication (read/write cells, execute functions)

**Python Example**:
```python
from aicalc_sdk import connect, get_value, set_value, run_function

# Connect to running AiCalc instance
connect()

# Read cell value
value = get_value("Sheet1!A1")
print(value)  # "42"

# Write cell value
set_value("Sheet1!B1", "Hello from Python")

# Execute AiCalc function
result = run_function("SUM", ["A1", "A2", "A3"])
print(result)  # "123"
```

**Custom Functions**:
```python
from aicalc_sdk import aicalc_function, CellObjectType

@aicalc_function(
    name="PYTHON_DOUBLE",
    description="Doubles a number using Python",
    category="Contrib",
    parameters=[("value", CellObjectType.Number)]
)
def double_value(value):
    return value * 2

# Function automatically available in AiCalc
```

**Pipe Protocol**:
```json
Request:  {"command": "get_value", "args": {"cell": "A1"}}
Response: {"success": true, "value": "42"}

Request:  {"command": "set_value", "args": {"cell": "B1", "value": "Hello"}}
Response: {"success": true}
```

---

## Build & Run

### Prerequisites

```powershell
# Check .NET SDK
dotnet --version  # Should be 8.0+

# Check Windows SDK
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots" | Select-Object KitsRoot10
```

---

### Build Commands

```powershell
# Quick build
dotnet build src/AiCalc.WinUI/AiCalc.WinUI.csproj

# Full rebuild
dotnet clean src/AiCalc.WinUI/AiCalc.WinUI.csproj
dotnet build src/AiCalc.WinUI/AiCalc.WinUI.csproj

# Release build
dotnet build src/AiCalc.WinUI/AiCalc.WinUI.csproj -c Release

# Build with self-contained runtime
dotnet publish src/AiCalc.WinUI/AiCalc.WinUI.csproj -c Debug -r win-x64 --self-contained /p:Platform=x64
```

---

### Run Commands

```powershell
# Quick run (Debug)
dotnet run --project src/AiCalc.WinUI/AiCalc.WinUI.csproj

# Run published executable
.\src\AiCalc.WinUI\bin\x64\Debug\net8.0-windows10.0.19041.0\win-x64\publish\AiCalc.WinUI.exe

# Launch scripts (convenience)
.\launch.ps1   # Fast launch (incremental build)
.\run.ps1      # Full rebuild + launch
```

---

### Test Commands

```powershell
# Run all tests
dotnet test tests/AiCalc.Tests/AiCalc.Tests.csproj

# Run with detailed output
dotnet test tests/AiCalc.Tests/AiCalc.Tests.csproj --logger "console;verbosity=detailed"

# Run specific test class
dotnet test tests/AiCalc.Tests/AiCalc.Tests.csproj --filter ClassName=CellAddressTests
```

---

### Visual Studio

```powershell
# Open solution
start AiCalc.sln

# Or open project directly
start src/AiCalc.WinUI/AiCalc.WinUI.csproj
```

**Debug Configuration**:
- Configuration: Debug
- Platform: x64
- Target Framework: net8.0-windows10.0.19041.0
- Debugger: Native (WinUI 3)

---

## Current Status (October 2025)

### Completed Features ✅

**Phase 1-2: Core Infrastructure**
- ✅ Models layer (14 classes)
- ✅ Cell object type system (31 types)
- ✅ Function registry (25+ functions)
- ✅ Function runner with formula parsing

**Phase 3: Multi-Threading**
- ✅ Dependency graph (DAG) with circular reference detection
- ✅ Topological sort for evaluation order
- ✅ Multi-threaded evaluation engine
- ✅ Cancellation token support
- ✅ Visual state system (7 states)
- ✅ Cell change visualization (borders, colors)
- ⏳ F9 recalculation (keyboard shortcut exists, exclude Manual mode pending)
- ⏳ Theme system (themes defined, Settings UI pending)

**Phase 4: AI Integration**
- ✅ AI service configuration with DPAPI encryption
- ✅ Multi-model support (Text, Vision, Image)
- ✅ Azure OpenAI client (245 lines)
- ✅ Ollama client (198 lines)
- ✅ Connection testing and validation
- ✅ AI function execution (9 AI functions)
- ✅ Formula autocomplete with type hints
- ✅ Contextual function suggestions
- ✅ AI response metadata tracking
- ⏳ Streaming responses (future)
- ⏳ Prompt templates (future)
- ⏳ Response caching (future)

**Phase 5: UI Polish**
- ✅ Keyboard navigation (8+ shortcuts)
- ✅ Context menus (13 operations)
- ✅ Cell visual states with themes
- ⏳ Enhanced formula bar (autocomplete working, syntax highlighting pending)

**Testing**
- ✅ 59 passing unit tests (xUnit)
- ✅ 100% test pass rate
- ✅ Coverage: Models, Services, DependencyGraph

---

### In Progress ⚠️

- **Phase 3**: Settings UI for thread count, per-service timeouts
- **Phase 4**: Task 13 (type-driven AI routing, cost estimation, batch operations)
- **Phase 5**: Task 11 (enhanced formula bar with syntax highlighting)

---

### Not Started ⏳

**Phase 6: Data Sources**
- Azure Blob, SQL Database integration
- Query builder UI
- Cell-level data binding

**Phase 7: Python SDK**
- Local environment detection
- Custom function discovery
- Cloud deployment (Azure Functions)

**Phase 8: Advanced Features**
- Cell Inspector tabs (Value, Formula, Notes, History)
- Automation workflow builder
- Macro recording

**Remaining Phases 1-2 Tasks**:
- Task 4: Excel-like navigation (arrows, Tab, Enter, Ctrl+Home/End, Page Up/Down)
- Task 5: Formula intellisense (parameter hints, cell picker, range notation)
- Task 6: Value state vs formula state toggle
- Task 7: Extract formula, spill operations, insert/displace

---

## Quick Reference Card

### Common Operations

**Create New Workbook**:
```csharp
var workbook = new WorkbookViewModel();
workbook.AddSheet();
var sheet = workbook.Sheets[0];
var cell = sheet.GetOrCreateCell(0, 0);  // A1
cell.Value = CellValue.FromText("Hello");
```

**Set Formula**:
```csharp
cell.Formula = "=SUM(A1:A10)";
await cell.EvaluateAsync();
```

**Configure AI Connection**:
```csharp
var connection = new WorkspaceConnection {
    Name = "My Azure OpenAI",
    Provider = "AzureOpenAI",
    Endpoint = "https://my-resource.openai.azure.com",
    ApiKeyEncrypted = CredentialService.Encrypt("sk-..."),
    Model = "gpt-4",
    IsActive = true
};
workbook.Settings.Connections.Add(connection);
```

**Execute AI Function**:
```csharp
var imageCell = sheet.GetOrCreateCell(0, 0);
imageCell.Value = CellValue.FromImage(@"C:\image.jpg");

var captionCell = sheet.GetOrCreateCell(0, 1);
captionCell.Formula = "=IMAGE_TO_CAPTION(A1)";
await captionCell.EvaluateAsync();
// Result: "A cat sitting on a couch"
```

**Save/Load**:
```csharp
await workbook.SaveAsync(@"C:\workbook.aicalc");
await workbook.LoadAsync(@"C:\workbook.aicalc");
```

---

### Key Files to Know

| File | Purpose | Lines |
|------|---------|-------|
| `WorkbookViewModel.cs` | Orchestration | 504 |
| `CellViewModel.cs` | Cell state & operations | 425 |
| `FunctionRegistry.cs` | Function catalog | 683 |
| `FunctionRunner.cs` | Formula execution | 322 |
| `EvaluationEngine.cs` | Multi-threaded evaluation | 314 |
| `DependencyGraph.cs` | DAG & topological sort | 238 |
| `AIServiceRegistry.cs` | AI client management | 134 |
| `MainWindow.xaml.cs` | Primary UI | 2037 |

---

### Architecture Patterns

**MVVM**: ViewModels expose properties/commands, UI binds via XAML  
**Service Locator**: Global `App.AIServices` for dependency injection  
**Repository**: `FunctionRegistry` manages function catalog  
**Observer**: `ObservableCollection`, `INotifyPropertyChanged` for reactive UI  
**Factory**: `CellValue.FromXxx()` static methods  
**Strategy**: `IAIServiceClient` interface for pluggable AI providers  
**Command**: `RelayCommand` for UI actions  

---

## Navigation Tips for Coding Agents

### Starting Points by Task

**"Add a new function"**  
→ `Services/FunctionRegistry.cs` → Find `RegisterBuiltIns()` → Add descriptor

**"Integrate new AI provider"**  
→ `Services/AI/` → Create new client class → Update `AIServiceRegistry.CreateClient()`

**"Add new cell type"**  
→ `Models/CellObjectType.cs` → Add enum → Update `CellValue` factory → Register functions

**"Fix formula parsing"**  
→ `Services/FunctionRunner.cs` → Look at `EvaluateAsync()` → Update regex pattern

**"Modify cell visual appearance"**  
→ `Converters/CellVisualStateToBrushConverter.cs` → Update color mapping

**"Add UI feature"**  
→ `MainWindow.xaml[.cs]` → Add XAML elements → Wire to ViewModel commands

**"Debug evaluation"**  
→ `Services/EvaluationEngine.cs` → Set breakpoints in `EvaluateCellAsync()`

**"Trace dependency issues"**  
→ `Services/DependencyGraph.cs` → Check `AddDependency()` and `TopologicalSort()`

**"Fix AI connection"**  
→ `ServiceConnectionDialog.xaml.cs` → Check `TestConnection_Click()` → Review `AIServiceRegistry`

**"Add test"**  
→ `tests/AiCalc.Tests/` → Create new test class → Follow xUnit patterns

---

## Glossary

**Cell Address**: Unique identifier for a cell (e.g., A1, Sheet2!B5)  
**Cell Value**: Typed value stored in a cell (number, text, image, etc.)  
**Cell Definition**: Complete cell state (value + formula + metadata)  
**Formula**: Expression starting with `=` that computes cell value  
**Function Descriptor**: Metadata for a function (name, parameters, handler)  
**Dependency Graph (DAG)**: Directed acyclic graph tracking cell references  
**Topological Sort**: Evaluation order respecting dependencies  
**Automation Mode**: When cell formula is evaluated (Manual, OnEdit, OnLoad, Continuous)  
**Visual State**: UI appearance state for cell (Normal, Updated, Calculating, Stale, Error)  
**Workspace Connection**: Configuration for AI service provider  
**DPAPI**: Windows Data Protection API for encrypting secrets  
**Spill Range**: Array formula result spanning multiple cells  
**Cell Object Type**: Classification of cell content (31 types)  

---

**End of Codebase Reference**  
For updates, see: `STATUS.md`, `tasks.md`, `docs/Phase*.md`
