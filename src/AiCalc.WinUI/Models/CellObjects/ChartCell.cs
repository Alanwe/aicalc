using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class ChartCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Chart;
    
    public string ChartType { get; set; }
    public string DataRange { get; set; }
    public string ConfigJson { get; set; }
    
    public override string? DisplayValue => $"📊 {ChartType} Chart";
    
    public ChartCell(string chartType, string dataRange, string configJson = "{}") 
        : base($"{chartType}|{dataRange}|{configJson}")
    {
        ChartType = chartType ?? "Bar";
        DataRange = dataRange ?? string.Empty;
        ConfigJson = configJson ?? "{}";
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(ChartType) && !string.IsNullOrWhiteSpace(DataRange);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "CHART_UPDATE";
        yield return "CHART_EXPORT_IMAGE";
        yield return "CHART_EXPORT_SVG";
        yield return "CHART_CONFIGURE";
        yield return "CHART_REFRESH";
    }
    
    public override ICellObject Clone() => new ChartCell(ChartType, DataRange, ConfigJson);
}
