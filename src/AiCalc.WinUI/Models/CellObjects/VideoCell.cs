using System;
using System.Collections.Generic;

namespace AiCalc.Models.CellObjects;

public class VideoCell : CellObjectBase
{
    public override CellObjectType ObjectType => CellObjectType.Video;
    
    public string FilePath { get; set; }
    public TimeSpan? Duration { get; set; }
    public string? Format { get; set; }
    
    public override string? DisplayValue => $"🎬 {System.IO.Path.GetFileName(FilePath)}";
    
    public VideoCell(string filePath) : base(filePath)
    {
        FilePath = filePath ?? string.Empty;
    }
    
    public override bool IsValid() => !string.IsNullOrWhiteSpace(FilePath);
    
    public override IEnumerable<string> GetAvailableOperations()
    {
        yield return "VIDEO_TO_AUDIO";
        yield return "VIDEO_EXTRACT_FRAMES";
        yield return "VIDEO_METADATA";
        yield return "VIDEO_THUMBNAIL";
        yield return "VIDEO_TRANSCRIBE";
        yield return "VIDEO_ANALYZE";
    }
    
    public override ICellObject Clone() => new VideoCell(FilePath)
    {
        Duration = Duration,
        Format = Format
    };
}
