using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class PdfCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Pdf;
    
    public string FilePath { get; set; }
    public int? PageCount { get; set; }
    
    public override string? DisplayValue => $"📕 {System.IO.Path.GetFileName(FilePath)} ({PageCount ?? 0} pages)";
    
    public PdfCell(string filePath) : base(filePath)
    {
        FilePath = filePath ?? string.Empty;
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(FilePath);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "PDF_TO_TEXT";
        yield return "PDF_TO_PAGES";
        yield return "PDF_PAGE_COUNT";
        yield return "PDF_EXTRACT_IMAGES";
        yield return "PDF_EXTRACT_TABLES";
        yield return "PDF_METADATA";
        yield return "PDF_MERGE";
        yield return "PDF_SPLIT";
        yield return "PDF_TO_IMAGE";
    }
    
    public override ICellObject Clone() => new PdfCell(FilePath)
    {
        PageCount = PageCount
    };
}
