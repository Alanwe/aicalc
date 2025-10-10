using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class PdfPageCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.PdfPage;
    
    public string PdfFilePath { get; set; }
    public int PageNumber { get; set; }
    
    public override string? DisplayValue => $"📄 Page {PageNumber} of {System.IO.Path.GetFileName(PdfFilePath)}";
    
    public PdfPageCell(string pdfFilePath, int pageNumber) : base($"{pdfFilePath}|{pageNumber}")
    {
        PdfFilePath = pdfFilePath ?? string.Empty;
        PageNumber = pageNumber;
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(PdfFilePath) && PageNumber > 0;
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "PDF_PAGE_TO_TEXT";
        yield return "PDF_PAGE_TO_IMAGE";
        yield return "PDF_PAGE_EXTRACT_TABLES";
        yield return "PDF_PAGE_EXTRACT_IMAGES";
        yield return "PDF_PAGE_OCR";
        yield return "PDF_PAGE_ANALYZE";
    }
    
    public override ICellObject Clone() => new PdfPageCell(PdfFilePath, PageNumber);
}
