using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class DirectoryCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Directory;
    
    public string DirectoryPath { get; set; }
    
    public override string? DisplayValue => $"📁 {System.IO.Path.GetFileName(DirectoryPath) ?? DirectoryPath}";
    
    public DirectoryCell(string directoryPath) : base(directoryPath)
    {
        DirectoryPath = directoryPath ?? string.Empty;
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(DirectoryPath);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "DIRECTORY_TO_TABLE";
        yield return "DIR_LIST";
        yield return "DIR_SIZE";
        yield return "DIR_COUNT";
        yield return "DIR_TREE";
        yield return "DIR_SEARCH";
        yield return "DIR_FILTER";
    }
    
    public override ICellObject Clone() => new DirectoryCell(DirectoryPath);
}
