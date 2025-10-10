using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class MarkdownCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Markdown;
    
    public string MarkdownText { get; set; }
    
    public override string? DisplayValue => $"📝 Markdown ({MarkdownText?.Length ?? 0} chars)";
    
    public MarkdownCell(string markdownText) : base(markdownText)
    {
        MarkdownText = markdownText ?? string.Empty;
    }
    
    public override bool IsValid() => MarkdownText != null;
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "MARKDOWN_TO_HTML";
        yield return "MARKDOWN_TO_TEXT";
        yield return "MARKDOWN_TO_PDF";
        yield return "MARKDOWN_PREVIEW";
        yield return "MARKDOWN_VALIDATE";
    }
    
    public override ICellObject Clone() => new MarkdownCell(MarkdownText);
}
