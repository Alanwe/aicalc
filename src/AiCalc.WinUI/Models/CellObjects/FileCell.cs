using System;
using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class FileCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.File;
    
    public string FilePath { get; set; }
    public long? FileSize { get; set; }
    public DateTime? LastModified { get; set; }
    
    public override string? DisplayValue => $"📄 {System.IO.Path.GetFileName(FilePath)}";
    
    public FileCell(string filePath) : base(filePath)
    {
        FilePath = filePath ?? string.Empty;
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(FilePath);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "FILE_READ";
        yield return "FILE_SIZE";
        yield return "FILE_EXTENSION";
        yield return "FILE_NAME";
        yield return "FILE_INFO";
        yield return "FILE_HASH";
        yield return "FILE_METADATA";
    }
    
    public override ICellObject Clone() => new FileCell(FilePath)
    {
        FileSize = FileSize,
        LastModified = LastModified
    };
}
