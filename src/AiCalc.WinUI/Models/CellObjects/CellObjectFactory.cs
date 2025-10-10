using System;

namespace AiCalc.Models.CellObjects;

/// <summary>
/// Factory for creating cell objects from values and types
/// </summary>
public static class CellObjectFactory
{
    public static ICellObject Create(CellObjectType type, string? value = null)
    {
        return type switch
        {
            CellObjectType.Empty => EmptyCell.Instance,
            CellObjectType.Number => new NumberCell(value ?? "0"),
            CellObjectType.Text => new TextCell(value ?? string.Empty),
            CellObjectType.Image => new ImageCell(value ?? string.Empty),
            CellObjectType.Video => new VideoCell(value ?? string.Empty),
            CellObjectType.Directory => new DirectoryCell(value ?? string.Empty),
            CellObjectType.File => new FileCell(value ?? string.Empty),
            CellObjectType.Table => new TableCell(value ?? "[]"),
            CellObjectType.Json => new JsonCell(value ?? "{}"),
            CellObjectType.Xml => new XmlCell(value ?? "<root/>"),
            CellObjectType.Markdown => new MarkdownCell(value ?? string.Empty),
            CellObjectType.Pdf => new PdfCell(value ?? string.Empty),
            CellObjectType.PdfPage => ParsePdfPage(value),
            CellObjectType.CodePython => CodeCell.CreatePython(value ?? string.Empty),
            CellObjectType.CodeCSharp => CodeCell.CreateCSharp(value ?? string.Empty),
            CellObjectType.CodeJavaScript => CodeCell.CreateJavaScript(value ?? string.Empty),
            CellObjectType.CodeTypeScript => CodeCell.CreateTypeScript(value ?? string.Empty),
            CellObjectType.CodeCss => CodeCell.CreateCss(value ?? string.Empty),
            CellObjectType.CodeHtml => CodeCell.CreateHtml(value ?? string.Empty),
            CellObjectType.CodeSql => CodeCell.CreateSql(value ?? string.Empty),
            CellObjectType.Chart => ParseChart(value),
            _ => EmptyCell.Instance
        };
    }
    
    public static ICellObject CreateFromCellValue(CellValue cellValue)
    {
        return Create(cellValue.ObjectType, cellValue.SerializedValue);
    }
    
    public static CellValue ToCellValue(ICellObject cellObject)
    {
        return new CellValue(
            cellObject.ObjectType,
            cellObject.SerializedValue,
            cellObject.DisplayValue
        );
    }
    
    private static ICellObject ParsePdfPage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return new PdfPageCell(string.Empty, 1);
            
        var parts = value.Split('|');
        if (parts.Length >= 2 && int.TryParse(parts[1], out var pageNum))
            return new PdfPageCell(parts[0], pageNum);
            
        return new PdfPageCell(value, 1);
    }
    
    private static ICellObject ParseChart(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return new ChartCell("Bar", string.Empty);
            
        var parts = value.Split('|');
        if (parts.Length >= 2)
            return new ChartCell(parts[0], parts[1], parts.Length > 2 ? parts[2] : "{}");
            
        return new ChartCell("Bar", value);
    }
}
