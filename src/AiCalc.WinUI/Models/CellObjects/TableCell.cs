using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class TableCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Table;
    
    public string JsonData { get; set; }
    public int? RowCount { get; set; }
    public int? ColumnCount { get; set; }
    
    public override string? DisplayValue => $"📊 Table ({RowCount ?? 0}×{ColumnCount ?? 0})";
    
    public TableCell(string jsonData) : base(jsonData)
    {
        JsonData = jsonData ?? "[]";
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(JsonData);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "TABLE_FILTER";
        yield return "TABLE_SORT";
        yield return "TABLE_JOIN";
        yield return "TABLE_AGGREGATE";
        yield return "TABLE_GROUP";
        yield return "TABLE_PIVOT";
        yield return "TABLE_TO_CSV";
        yield return "TABLE_TO_EXCEL";
        yield return "TABLE_COLUMN";
        yield return "TABLE_ROW";
    }
    
    public override ICellObject Clone() => new TableCell(JsonData)
    {
        RowCount = RowCount,
        ColumnCount = ColumnCount
    };
}
