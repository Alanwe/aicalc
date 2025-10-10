using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AiCalc.Models.CellObjects;

/// <summary>
/// Base interface for all cell object types
/// </summary>
public interface ICellObject
{
    /// <summary>
    /// The type of this cell object
    /// </summary>
    CellObjectType ObjectType { get; }
    
    /// <summary>
    /// Serialized representation for persistence
    /// </summary>
    string? SerializedValue { get; set; }
    
    /// <summary>
    /// Display representation for UI
    /// </summary>
    string? DisplayValue { get; }
    
    /// <summary>
    /// Validates the cell object's data
    /// </summary>
    bool IsValid();
    
    /// <summary>
    /// Gets available operations for this cell type
    /// </summary>
    IEnumerable<string> GetAvailableOperations();
    
    /// <summary>
    /// Clones this cell object
    /// </summary>
    ICellObject Clone();
}

/// <summary>
/// Base abstract class for cell objects with common functionality
/// </summary>
public abstract class CellObjectBase : ICellObject
{
    public abstract CellObjectType ObjectType { get; }
    public string? SerializedValue { get; set; }
    public abstract string? DisplayValue { get; }
    
    public virtual bool IsValid() => true;
    
    public abstract IEnumerable<string> GetAvailableOperations();
    
    public abstract ICellObject Clone();
    
    protected CellObjectBase(string? serializedValue = null)
    {
        SerializedValue = serializedValue;
    }
}
