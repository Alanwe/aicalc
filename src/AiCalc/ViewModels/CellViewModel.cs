using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AiCalc.Models;
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

    public string DisplayLabel => Address.ToString();

    public string ObjectTypeGlyph => Value.ObjectType switch
    {
        CellObjectType.Number => "",
        CellObjectType.Text => "",
        CellObjectType.Image => "",
        CellObjectType.Video => "",
        CellObjectType.Audio => "",
        CellObjectType.Directory => "",
        CellObjectType.File => "",
        CellObjectType.Table => "",
        CellObjectType.Script => "",
        CellObjectType.Markdown => "",
        CellObjectType.Json => "{}",
        CellObjectType.Link => "",
        CellObjectType.Error => "⚠",
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
            _workbook.IsBusy = true;
            _workbook.StatusMessage = $"Evaluating {Address}...";
            var result = await _workbook.FunctionRunner.EvaluateAsync(this, Formula);
            if (result is not null)
            {
                Value = result.Value;
            }
        }
        catch (Exception ex)
        {
            Value = new CellValue(CellObjectType.Error, ex.Message, ex.Message);
        }
        finally
        {
            _workbook.StatusMessage = null;
            _workbook.IsBusy = false;
        }
    }

    [RelayCommand]
    private Task RunAsync() => EvaluateAsync();

    partial void OnValueChanged(CellValue value)
    {
        OnPropertyChanged(nameof(DisplayValue));
        OnPropertyChanged(nameof(ObjectTypeGlyph));
        OnPropertyChanged(nameof(RawValue));
    }
}

