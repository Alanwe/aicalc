using System;
using System.Collections.Generic;
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

    public string? RawValue
    {
        get => Value.SerializedValue;
        set
        {
            var objectType = Value.ObjectType == CellObjectType.Empty && !string.IsNullOrWhiteSpace(value) ? CellObjectType.Text : Value.ObjectType;
            var display = string.IsNullOrWhiteSpace(Value.DisplayValue) ? value : Value.DisplayValue;
            Value = new CellValue(objectType, value, display);
        }
    }

    public bool HasFormula => !string.IsNullOrWhiteSpace(Formula);

    public async Task EvaluateAsync()
    {
        if (string.IsNullOrWhiteSpace(Formula))
        {
            return;
        }

        try
        {
            MarkAsCalculating();
            _workbook.IsBusy = true;
            _workbook.StatusMessage = $"Evaluating {Address}...";
            
            var result = await _workbook.FunctionRunner.EvaluateAsync(this, Formula);
            if (result is not null)
            {
                Value = result.Value;
                MarkAsUpdated();
            }
        }
        catch (Exception ex)
        {
            Value = new CellValue(CellObjectType.Error, ex.Message, ex.Message);
            VisualState = CellVisualState.Error;
            IsCalculating = false;
        }
        finally
        {
            _workbook.StatusMessage = null;
            _workbook.IsBusy = false;
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
}

