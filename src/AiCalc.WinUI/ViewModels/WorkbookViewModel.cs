using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.Services;
using AiCalc.Services.AI;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AiCalc.ViewModels;

public partial class WorkbookViewModel : BaseViewModel
{
    private readonly JsonSerializerOptions _serializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
    };

    private readonly DependencyGraph _dependencyGraph;
    private readonly EvaluationEngine _evaluationEngine;
    private readonly UndoRedoManager _undoRedoManager;

    public WorkbookViewModel()
    {
        FunctionRegistry = new FunctionRegistry();
        FunctionRunner = new FunctionRunner(FunctionRegistry);
        _dependencyGraph = new DependencyGraph();
        _evaluationEngine = new EvaluationEngine(
            _dependencyGraph, 
            FunctionRunner,
            Settings.MaxEvaluationThreads,
            Settings.DefaultEvaluationTimeoutSeconds);
        _undoRedoManager = new UndoRedoManager(maxUndoLevels: 50);
        _undoRedoManager.StacksChanged += (s, e) =>
        {
            UndoCommand.NotifyCanExecuteChanged();
            RedoCommand.NotifyCanExecuteChanged();
        };
        
        Sheets = new ObservableCollection<SheetViewModel>();
        AttachSettings(Settings);
        AddSheet();
        SelectedSheet = Sheets.FirstOrDefault();
    }

    public FunctionRegistry FunctionRegistry { get; }

    public FunctionRunner FunctionRunner { get; }

    public DependencyGraph DependencyGraph => _dependencyGraph;

    public EvaluationEngine EvaluationEngine => _evaluationEngine;

    public UndoRedoManager UndoRedoManager => _undoRedoManager;

    /// <summary>
    /// Record a cell change for undo/redo (Phase 5)
    /// </summary>
    public void RecordCellChange(CellChangeAction action)
    {
        _undoRedoManager.RecordAction(action);
    }

    public ObservableCollection<SheetViewModel> Sheets { get; }

    public Array AutomationModes { get; } = Enum.GetValues(typeof(CellAutomationMode));

    [ObservableProperty]
    private SheetViewModel? _selectedSheet;

    [ObservableProperty]
    private WorkbookSettings _settings = new();

    [ObservableProperty]
    private string _title = "Untitled Workbook";

    [ObservableProperty]
    private CellViewModel? _activeCell;

    [ObservableProperty]
    private string _statusMessage = "Ready";

    public bool HasActiveCell => ActiveCell is not null;

    partial void OnActiveCellChanged(CellViewModel? oldValue, CellViewModel? newValue)
    {
        foreach (var cell in Sheets.SelectMany(s => s.Cells))
        {
            cell.IsSelected = cell == newValue;
        }

        OnPropertyChanged(nameof(HasActiveCell));
    }

    public void AddSheet()
    {
        var index = Sheets.Count + 1;
        var sheet = new SheetViewModel(this, $"Sheet{index}", 20, 12);
        Sheets.Add(sheet);
    }

    public SheetViewModel? GetSheet(string name) => Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

    public CellViewModel? GetCell(CellAddress address)
    {
        var sheet = GetSheet(address.SheetName);
        return sheet?.GetCell(address.Row, address.Column);
    }

    [RelayCommand]
    private void NewSheet()
    {
        AddSheet();
        SelectedSheet = Sheets.Last();
    }

    [RelayCommand]
    private void SelectCell(CellViewModel cell)
    {
        ActiveCell = cell;
    }

    [RelayCommand]
    private async Task EvaluateSheetAsync()
    {
        if (SelectedSheet is null)
        {
            return;
        }

        await SelectedSheet.EvaluateAllAsync();
    }

    [RelayCommand]
    private async Task EvaluateWorkbookAsync()
    {
        // Recalculate All (F9) - Skip Manual cells per Task 10
        var cellsToEvaluate = Sheets
            .SelectMany(s => s.Cells)
            .Where(c => c.HasFormula && c.AutomationMode != CellAutomationMode.Manual)
            .ToDictionary(c => c.Address, c => c);

        if (cellsToEvaluate.Count == 0)
        {
            StatusMessage = "No cells to evaluate";
            return;
        }

        StatusMessage = $"Evaluating {cellsToEvaluate.Count} cells...";
        
        try
        {
            var result = await _evaluationEngine.EvaluateAllAsync(cellsToEvaluate);
            
            if (result.CellsEvaluated > 0 && result.CellsFailed == 0)
            {
                StatusMessage = $"✅ Successfully evaluated {result.CellsEvaluated} cells in {result.Duration.TotalSeconds:F2}s";
            }
            else if (result.CellsEvaluated > 0 && result.CellsFailed > 0)
            {
                StatusMessage = $"⚠️ Evaluated {result.CellsEvaluated}/{cellsToEvaluate.Count} cells, {result.CellsFailed} errors";
            }
            else
            {
                StatusMessage = "Ready";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"❌ Evaluation error: {ex.Message}";
        }
    }

    /// <summary>
    /// Updates evaluation engine settings from WorkbookSettings
    /// </summary>
    public void UpdateEvaluationSettings()
    {
        _evaluationEngine.MaxDegreeOfParallelism = Settings.MaxEvaluationThreads;
        _evaluationEngine.DefaultTimeoutSeconds = Settings.DefaultEvaluationTimeoutSeconds;
    }

    [RelayCommand]
    private async Task SaveAsync()
    {
        var definition = ToDefinition();
        var fileName = Title.Replace(' ', '_') + ".aicalc";
        var json = JsonSerializer.Serialize(definition, _serializerOptions);
        await File.WriteAllTextAsync(fileName, json);
        StatusMessage = $"Saved to {fileName}";
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        var fileName = Title.Replace(' ', '_') + ".aicalc";
        if (!File.Exists(fileName))
        {
            StatusMessage = $"File {fileName} not found";
            return;
        }

        var json = await File.ReadAllTextAsync(fileName);
        var definition = JsonSerializer.Deserialize<WorkbookDefinition>(json, _serializerOptions);
        if (definition is null)
        {
            return;
        }

        ApplyDefinition(definition);
        StatusMessage = $"Loaded {fileName}";
    }

    [RelayCommand]
    private void ClearSelection()
    {
        foreach (var cell in Sheets.SelectMany(s => s.Cells))
        {
            cell.IsSelected = false;
        }

        ActiveCell = null;
    }

    public WorkbookDefinition ToDefinition()
    {
        var definition = new WorkbookDefinition
        {
            Title = Title,
            Settings = Settings
        };

        foreach (var sheet in Sheets)
        {
            var sheetDef = new SheetDefinition
            {
                Name = sheet.Name,
                RowCount = sheet.Rows.Count,
                ColumnCount = sheet.ColumnCount
            };

            foreach (var cell in sheet.Cells)
            {
                if (cell.Value.ObjectType == CellObjectType.Empty && string.IsNullOrWhiteSpace(cell.Formula))
                {
                    continue;
                }

                sheetDef.Cells.Add(new CellDefinition
                {
                    Address = cell.Address.ToString(),
                    Formula = cell.Formula,
                    Value = cell.Value,
                    AutomationMode = cell.AutomationMode,
                    Notes = cell.Notes,
                    Format = cell.Format.Clone(),
                    History = new ObservableCollection<CellHistoryEntry>(cell.History.Select(h => h.Clone()))
                });
            }

            definition.Sheets.Add(sheetDef);
        }

        return definition;
    }

    private void ApplyDefinition(WorkbookDefinition definition)
    {
        Sheets.Clear();
        Title = definition.Title;
        Settings = definition.Settings;

        foreach (var sheet in definition.Sheets)
        {
            var sheetVm = new SheetViewModel(this, sheet.Name, sheet.RowCount, sheet.ColumnCount);
            foreach (var cellDef in sheet.Cells)
            {
                if (!CellAddress.TryParse(cellDef.Address, sheet.Name, out var address))
                {
                    continue;
                }

                var cell = sheetVm.GetCell(address.Row, address.Column);
                if (cell is null)
                {
                    continue;
                }

                using (cell.SuppressHistory())
                {
                    cell.LoadFromDefinition(cellDef);
                }
            }

            Sheets.Add(sheetVm);
        }

        SelectedSheet = Sheets.FirstOrDefault();
    }

    partial void OnSettingsChanging(WorkbookSettings value)
    {
        if (value != null)
        {
            value.Connections.CollectionChanged -= OnConnectionsCollectionChanged;
        }
    }

    partial void OnSettingsChanged(WorkbookSettings value)
    {
        if (value != null)
        {
            AttachSettings(value);
        }
    }

    private void AttachSettings(WorkbookSettings settings)
    {
        settings.Connections.CollectionChanged -= OnConnectionsCollectionChanged;
        settings.Connections.CollectionChanged += OnConnectionsCollectionChanged;
        SynchronizeConnections(settings);
    }

    private void SynchronizeConnections(WorkbookSettings settings)
    {
        App.AIServices.Clear();
        foreach (var connection in settings.Connections)
        {
            EnsureEncryptedApiKey(connection);
            if (connection.IsActive)
            {
                App.AIServices.RegisterConnection(connection);
            }
        }
    }

    private void OnConnectionsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        switch (e.Action)
        {
            case NotifyCollectionChangedAction.Add:
                if (e.NewItems != null)
                {
                    foreach (WorkspaceConnection connection in e.NewItems)
                    {
                        EnsureEncryptedApiKey(connection);
                        if (connection.IsActive)
                        {
                            App.AIServices.RegisterConnection(connection);
                        }
                    }
                }
                break;

            case NotifyCollectionChangedAction.Remove:
                if (e.OldItems != null)
                {
                    foreach (WorkspaceConnection connection in e.OldItems)
                    {
                        App.AIServices.RemoveConnection(connection.Id);
                    }
                }
                break;

            case NotifyCollectionChangedAction.Replace:
                if (e.OldItems != null)
                {
                    foreach (WorkspaceConnection connection in e.OldItems)
                    {
                        App.AIServices.RemoveConnection(connection.Id);
                    }
                }

                if (e.NewItems != null)
                {
                    foreach (WorkspaceConnection connection in e.NewItems)
                    {
                        EnsureEncryptedApiKey(connection);
                        if (connection.IsActive)
                        {
                            App.AIServices.RegisterConnection(connection);
                        }
                    }
                }
                break;

            case NotifyCollectionChangedAction.Reset:
                App.AIServices.Clear();
                foreach (var connection in Settings.Connections)
                {
                    EnsureEncryptedApiKey(connection);
                    if (connection.IsActive)
                    {
                        App.AIServices.RegisterConnection(connection);
                    }
                }
                break;

            case NotifyCollectionChangedAction.Move:
                // No registry changes required when items reorder.
                break;
        }
    }

    private static void EnsureEncryptedApiKey(WorkspaceConnection connection)
    {
        if (string.IsNullOrWhiteSpace(connection.ApiKey))
        {
            return;
        }

        if (!CredentialService.IsEncrypted(connection.ApiKey))
        {
            connection.ApiKey = CredentialService.Encrypt(connection.ApiKey);
        }
    }

    /// <summary>
    /// Undo the last cell change (Phase 5)
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanUndo))]
    private void Undo()
    {
        var action = _undoRedoManager.Undo();
        if (action == null) return;

        var cell = FindCell(action.Address);
        if (cell == null) return;

        ApplyActionReverse(cell, action);
        StatusMessage = $"↶ Undo: {action.Description}";
    }

    private bool CanUndo() => _undoRedoManager.CanUndo;

    /// <summary>
    /// Redo the last undone cell change (Phase 5)
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanRedo))]
    private void Redo()
    {
        var action = _undoRedoManager.Redo();
        if (action == null) return;

        var cell = FindCell(action.Address);
        if (cell == null) return;

        ApplyActionForward(cell, action);
        StatusMessage = $"↷ Redo: {action.Description}";
    }

    private bool CanRedo() => _undoRedoManager.CanRedo;

    private void ApplyActionReverse(CellViewModel cell, CellChangeAction action)
    {
        if (action.OldFormula != null || action.NewFormula != null)
        {
            cell.Formula = action.OldFormula ?? "";
        }
        
        if (action.OldFormat != null || action.NewFormat != null)
        {
            cell.Format = action.OldFormat!;
        }
        
        if (action.OldAutomationMode != action.NewAutomationMode)
        {
            cell.AutomationMode = action.OldAutomationMode;
        }
    }

    private void ApplyActionForward(CellViewModel cell, CellChangeAction action)
    {
        if (action.OldFormula != null || action.NewFormula != null)
        {
            cell.Formula = action.NewFormula ?? "";
        }
        
        if (action.OldFormat != null || action.NewFormat != null)
        {
            cell.Format = action.NewFormat!;
        }
        
        if (action.OldAutomationMode != action.NewAutomationMode)
        {
            cell.AutomationMode = action.NewAutomationMode;
        }
    }

    private CellViewModel? FindCell(CellAddress address)
    {
        var sheet = Sheets.FirstOrDefault(s => s.Name == address.SheetName);
        return sheet?.Cells.FirstOrDefault(c => c.Address.Row == address.Row && c.Address.Column == address.Column);
    }
}
