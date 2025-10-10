using System;
using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class ImageCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Image;
    
    public string FilePath { get; set; }
    public int? Width { get; set; }
    public int? Height { get; set; }
    public string? Format { get; set; }
    
    public override string? DisplayValue => $"Image: {System.IO.Path.GetFileName(FilePath)}";
    
    public ImageCell(string filePath) : base(filePath)
    {
        FilePath = filePath ?? string.Empty;
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(FilePath);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "IMAGE_TO_TEXT";
        yield return "IMAGE_TO_CAPTION";
        yield return "IMAGE_RESIZE";
        yield return "IMAGE_CROP";
        yield return "IMAGE_CONVERT";
        yield return "IMAGE_METADATA";
        yield return "IMAGE_OCR";
        yield return "IMAGE_ANALYZE";
    }
    
    public override ICellObject Clone() => new ImageCell(FilePath)
    {
        Width = Width,
        Height = Height,
        Format = Format
    };
}
