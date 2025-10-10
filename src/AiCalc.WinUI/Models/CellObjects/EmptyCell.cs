using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class EmptyCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Empty;
    public override string? DisplayValue => string.Empty;
    
    public EmptyCell() : base(null) { }
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield break; // No operations for empty cells
    }
    
    public override ICellObject Clone() => new EmptyCell();
    
    public static EmptyCell Instance { get; } = new EmptyCell();
}
