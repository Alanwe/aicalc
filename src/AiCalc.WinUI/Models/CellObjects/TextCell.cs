using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class TextCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Text;
    
    public string Text { get; set; }
    
    public override string? DisplayValue => Text;
    
    public TextCell(string text) : base(text)
    {
        Text = text ?? string.Empty;
    }
    
    public override bool IsValid() => Text != null;
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "CONCAT";
        yield return "UPPER";
        yield return "LOWER";
        yield return "TRIM";
        yield return "LEN";
        yield return "REPLACE";
        yield return "SPLIT";
        yield return "SUBSTRING";
        yield return "TEXT_TO_IMAGE";
        yield return "TEXT_TO_SPEECH";
    }
    
    public override ICellObject Clone() => new TextCell(Text);
}
