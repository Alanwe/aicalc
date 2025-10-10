using System;
using System.Collections.Generic;
using System.Globalization;

namespace AiCalc.Models.CellObjects;

public class NumberCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Number;
    
    public double Value { get; set; }
    
    public override string? DisplayValue => Value.ToString(CultureInfo.CurrentCulture);
    
    public NumberCell(double value) : base(value.ToString(CultureInfo.InvariantCulture))
    {
        Value = value;
    }
    
    public NumberCell(string serializedValue) : base(serializedValue)
    {
        if (double.TryParse(serializedValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
        {
            Value = value;
        }
    }
    
    public override bool IsValid() => !double.IsNaN(Value) && !double.IsInfinity(Value);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "SUM";
        yield return "AVERAGE";
        yield return "MIN";
        yield return "MAX";
        yield return "ROUND";
        yield return "ABS";
        yield return "SQRT";
        yield return "POWER";
    }
    
    public override ICellObject Clone() => new NumberCell(Value);
}
