using System.Text.Json.Serialization;

namespace AiCalc.Models;

public record CellValue
{
    [JsonConstructor]
    public CellValue(CellObjectType objectType, string? serializedValue = null, string? displayValue = null)
    {
        ObjectType = objectType;
        SerializedValue = serializedValue;
        DisplayValue = displayValue;
    }

    public CellValue()
    {
    }

    public CellObjectType ObjectType { get; set; }

    public string? SerializedValue { get; set; }

    public string? DisplayValue { get; set; }

    public static CellValue Empty { get; } = new(CellObjectType.Empty, null, string.Empty);
}
