using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.Models.CellObjects;
using AiCalc.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AiCalc.ViewModels;

public partial class CellViewModel : ObservableObject
{
    private readonly WorkbookViewModel _workbook;
    private readonly ObservableCollection<CellHistoryEntry> _history = new();
    private CellFormat _format = CellFormat.Default;
    private int _historySuppressionDepth;
    private string? _formulaBeforeChange;

    public CellViewModel(WorkbookViewModel workbook, SheetViewModel sheet, int row, int column)
    {
        _workbook = workbook;
        Sheet = sheet;
        Row = row;
        Column = column;
        Address = new CellAddress(sheet.Name, row, column);
    }

    public SheetViewModel Sheet { get; }

    public int Row { get; }

    public int Column { get; }

    public CellAddress Address { get; }

    [ObservableProperty]
    private string? _formula;

    [ObservableProperty]
    private CellValue _value = CellValue.Empty;

    [ObservableProperty]
    private CellAutomationMode _automationMode = CellAutomationMode.Manual;

    [ObservableProperty]
    private string? _notes;

    [ObservableProperty]
    private bool _isSelected;
    
    [ObservableProperty]
    private CellVisualState _visualState = CellVisualState.Normal;
    
    [ObservableProperty]
    private bool _isCalculating;
    
    [ObservableProperty]
    private DateTime? _lastUpdated;
    
    [ObservableProperty]
    private bool _isStale;
    
    public ObservableCollection<CellHistoryEntry> History => _history;

    public bool HasHistory => _history.Count > 0;

    public CellFormat Format
    {
        get => _format;
        set
        {
            var normalized = CellFormat.From(value);
            if (SetProperty(ref _format, normalized))
            {
                OnPropertyChanged(nameof(FormatBackground));
                OnPropertyChanged(nameof(FormatForeground));
                OnPropertyChanged(nameof(FormatBorder));
                OnPropertyChanged(nameof(FormatFontSize));
            }
        }
    }

    public string FormatBackground => Format.Background;

    public string FormatForeground => Format.Foreground;

    public string FormatBorder => Format.BorderBrush;

    public double FormatFontSize => Format.FontSize;

    private ICellObject? _cellObject;
    
    public ICellObject CellObject
    {
        get => _cellObject ?? EmptyCell.Instance;
        set
        {
            _cellObject = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(AvailableOperations));
        }
    }
    
    public IEnumerable<string> AvailableOperations => CellObject.GetAvailableOperations();

    public string DisplayLabel => Address.ToString();

    public string ObjectTypeGlyph => Value.ObjectType switch
    {
        CellObjectType.Number => "🔢",
        CellObjectType.Text => "📝",
        CellObjectType.Image => "🖼️",
        CellObjectType.Video => "🎬",
        CellObjectType.Audio => "🔊",
        CellObjectType.Directory => "📁",
        CellObjectType.File => "📄",
        CellObjectType.Table => "📊",
        CellObjectType.Script => "📜",
        CellObjectType.Markdown => "📝",
        CellObjectType.Json => "{}",
        CellObjectType.Xml => "<>",
        CellObjectType.Link => "🔗",
        CellObjectType.Error => "⚠",
        CellObjectType.Pdf => "📕",
        CellObjectType.PdfPage => "📄",
        CellObjectType.CodePython => "🐍",
        CellObjectType.CodeCSharp => "#️⃣",
        CellObjectType.CodeJavaScript => "JS",
        CellObjectType.CodeTypeScript => "TS",
        CellObjectType.CodeCss => "🎨",
        CellObjectType.CodeHtml => "🌐",
        CellObjectType.CodeSql => "🗄️",
        CellObjectType.Chart => "📈",
        CellObjectType.ChartImage => "📊",
        _ => "○"
    };

    public string DisplayValue => string.IsNullOrWhiteSpace(Value.DisplayValue) ? Value.SerializedValue ?? string.Empty : Value.DisplayValue;

    private bool IsHistoryActive => _historySuppressionDepth == 0;

    public IDisposable SuppressHistory()
    {
        _historySuppressionDepth++;
        return new HistoryScope(this);
    }

    private void ResumeHistory()
    {
        if (_historySuppressionDepth > 0)
        {
            _historySuppressionDepth--;
        }
    }

    private void AppendHistory(CellValue oldValue, CellValue newValue, string? oldFormula, string? newFormula, string reason)
    {
        if (!IsHistoryActive)
        {
            return;
        }

        var valueChanged = oldValue.ObjectType != newValue.ObjectType ||
                           !string.Equals(oldValue.SerializedValue, newValue.SerializedValue, StringComparison.Ordinal) ||
                           !string.Equals(oldValue.DisplayValue, newValue.DisplayValue, StringComparison.Ordinal);

        var formulaChanged = !string.Equals(oldFormula, newFormula, StringComparison.Ordinal);

        if (!valueChanged && !formulaChanged)
        {
            return;
        }

        var entry = new CellHistoryEntry
        {
            Timestamp = DateTime.UtcNow,
            OldValue = oldValue,
            NewValue = newValue,
            OldFormula = oldFormula,
            NewFormula = newFormula,
            Notes = reason
        };

        _history.Add(entry);

        var max = Math.Max(5, _workbook.Settings.MaxHistoryEntries);
        while (_history.Count > max)
        {
            _history.RemoveAt(0);
        }

        OnPropertyChanged(nameof(History));

        // Record for undo/redo (Phase 5)
        if (valueChanged || formulaChanged)
        {
            var action = CellChangeAction.ForValueChange(
                Address,
                oldValue.DisplayValue,
                newValue.DisplayValue,
                oldFormula,
                newFormula,
                AutomationMode,
                AutomationMode,
                reason);
            _workbook.RecordCellChange(action);
        }
        OnPropertyChanged(nameof(HasHistory));
    }

    public void ReplaceHistory(IEnumerable<CellHistoryEntry> entries)
    {
        using (SuppressHistory())
        {
            _history.Clear();
            foreach (var entry in entries)
            {
                _history.Add(entry.Clone());
            }
        }

        OnPropertyChanged(nameof(History));
        OnPropertyChanged(nameof(HasHistory));
    }

    public void ApplyEvaluationResult(FunctionExecutionResult result, string reason = "Evaluated")
    {
        var oldValue = Value;
        using (SuppressHistory())
        {
            Value = result.Value;
        }

        var historyReason = reason;
        if (!string.IsNullOrWhiteSpace(result.Diagnostics))
        {
            historyReason = string.IsNullOrWhiteSpace(reason)
                ? result.Diagnostics
                : $"{reason} • {result.Diagnostics}";
        }

        AppendHistory(oldValue, result.Value, Formula, Formula, historyReason);
        MarkAsUpdated();
    }

    internal void CopyFrom(CellViewModel source)
    {
        using (SuppressHistory())
        {
            Format = source.Format.Clone();
            AutomationMode = source.AutomationMode;
            Notes = source.Notes;
            Value = source.Value;
            Formula = source.Formula;
            _history.Clear();
            foreach (var entry in source.History)
            {
                _history.Add(entry.Clone());
            }
        }

        OnPropertyChanged(nameof(History));
        OnPropertyChanged(nameof(HasHistory));
    }

    internal void LoadFromDefinition(CellDefinition definition)
    {
        using (SuppressHistory())
        {
            Format = definition.Format ?? CellFormat.Default;
            var persistedHistory = definition.History as IEnumerable<CellHistoryEntry> ?? Array.Empty<CellHistoryEntry>();
            ReplaceHistory(persistedHistory);
            AutomationMode = definition.AutomationMode;
            Notes = definition.Notes;
            Value = definition.Value ?? CellValue.Empty;
            Formula = definition.Formula;
        }

        _cellObject = CellObjectFactory.CreateFromCellValue(Value);
        OnPropertyChanged(nameof(DisplayValue));
        OnPropertyChanged(nameof(ObjectTypeGlyph));
        OnPropertyChanged(nameof(RawValue));
        OnPropertyChanged(nameof(CellObject));
        OnPropertyChanged(nameof(AvailableOperations));
    }

    public string? RawValue
    {
        get => Value.SerializedValue;
        set
        {
            var incoming = value ?? string.Empty;
            var currentSerialized = Value.SerializedValue ?? string.Empty;

            if (string.Equals(incoming, currentSerialized, StringComparison.Ordinal))
            {
                return;
            }

            string processed = incoming;
            CellObjectType objectType = Value.ObjectType;

            // If starts with ', treat as string and remove '
            if (processed.StartsWith("'"))
            {
                processed = processed.Substring(1);
                objectType = CellObjectType.Text;
            }
            else if (double.TryParse(processed, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                // If it's a valid number, classify as Number
                objectType = CellObjectType.Number;
            }
            else
            {
                objectType = CellObjectType.Text;
            }

            var display = processed;
            var oldValue = Value;
            var newValue = new CellValue(objectType, processed, display);

            AppendHistory(oldValue, newValue, Formula, Formula, "Value edited");

            using (SuppressHistory())
            {
                Value = newValue;
            }

            MarkAsUpdated();
        }
    }

    public bool HasFormula => !string.IsNullOrWhiteSpace(Formula);

    public async Task<FunctionExecutionResult?> EvaluateAsync()
    {
        if (string.IsNullOrWhiteSpace(Formula))
        {
            return null;
        }

        try
        {
            MarkAsCalculating();
            _workbook.IsBusy = true;
            _workbook.StatusMessage = $"Evaluating {Address}...";
            
            var result = await _workbook.FunctionRunner.EvaluateAsync(this, Formula);
            return result;
        }
        catch (Exception ex)
        {
            Value = new CellValue(CellObjectType.Error, ex.Message, ex.Message);
            VisualState = CellVisualState.Error;
            IsCalculating = false;
            return null;
        }
        finally
        {
            _workbook.StatusMessage = "Ready";
            _workbook.IsBusy = false;
            IsCalculating = false;
            if (!IsStale && VisualState != CellVisualState.Error)
            {
                VisualState = CellVisualState.Normal;
            }
        }
    }

    [RelayCommand]
    private Task RunAsync() => EvaluateAsync();
    
    /// <summary>
    /// Marks this cell as stale (needs recalculation)
    /// </summary>
    public void MarkAsStale()
    {
        IsStale = true;
        if (VisualState == CellVisualState.Normal)
        {
            VisualState = CellVisualState.Stale;
        }
    }
    
    /// <summary>
    /// Marks this cell as calculating
    /// </summary>
    public void MarkAsCalculating()
    {
        IsCalculating = true;
        VisualState = CellVisualState.Calculating;
    }
    
    /// <summary>
    /// Marks this cell as updated and triggers visual flash effect
    /// </summary>
    public async void MarkAsUpdated()
    {
        IsStale = false;
        IsCalculating = false;
        LastUpdated = DateTime.Now;
        VisualState = CellVisualState.JustUpdated;
        
        // Mark workbook as dirty for autosave (Phase 6)
        _workbook.MarkDirty();
        
        // Flash effect: return to normal after 2 seconds
        await Task.Delay(2000);
        if (VisualState == CellVisualState.JustUpdated)
        {
            VisualState = CellVisualState.Normal;
        }
    }
    
    /// <summary>
    /// Clears all visual states
    /// </summary>
    public void ClearVisualState()
    {
        IsStale = false;
        IsCalculating = false;
        if (VisualState != CellVisualState.Error)
        {
            VisualState = CellVisualState.Normal;
        }
    }

    partial void OnValueChanged(CellValue value)
    {
        // Update the cell object when value changes
        _cellObject = CellObjectFactory.CreateFromCellValue(value);
        
        OnPropertyChanged(nameof(DisplayValue));
        OnPropertyChanged(nameof(ObjectTypeGlyph));
        OnPropertyChanged(nameof(RawValue));
        OnPropertyChanged(nameof(CellObject));
        OnPropertyChanged(nameof(AvailableOperations));
    }

    partial void OnFormulaChanging(string? value)
    {
    _formulaBeforeChange = Formula;
    }

    partial void OnFormulaChanged(string? value)
    {
        if (_formulaBeforeChange != value)
        {
            AppendHistory(Value, Value, _formulaBeforeChange, value, "Formula edited");
            Sheet.UpdateCellDependencies(this);
            MarkAsStale();
        }

        _formulaBeforeChange = null;
    }

    private sealed class HistoryScope : IDisposable
    {
        private readonly CellViewModel _owner;
        private bool _disposed;

        public HistoryScope(CellViewModel owner)
        {
            _owner = owner;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _owner.ResumeHistory();
        }
    }
}

