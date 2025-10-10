using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class JsonCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Json;
    
    public string JsonText { get; set; }
    
    public override string? DisplayValue => $"{{}} JSON ({JsonText?.Length ?? 0} chars)";
    
    public JsonCell(string jsonText) : base(jsonText)
    {
        JsonText = jsonText ?? "{}";
    }
    
    public override bool IsValid()
    {
        try
        {
            System.Text.Json.JsonDocument.Parse(JsonText);
            return true;
        }
        catch
        {
            return false;
        }
    }
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "JSON_VALIDATE";
        yield return "JSON_FORMAT";
        yield return "JSON_MINIFY";
        yield return "JSON_QUERY";
        yield return "JSON_PATH";
        yield return "JSON_TO_TABLE";
        yield return "JSON_MERGE";
    }
    
    public override ICellObject Clone() => new JsonCell(JsonText);
}
