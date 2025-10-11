using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AiCalc.Models;

namespace AiCalc.Services;

/// <summary>
/// Represents a node in the dependency graph
/// </summary>
public class DependencyNode
{
    public CellAddress Address { get; set; }
    public HashSet<CellAddress> Dependencies { get; set; } = new();
    public HashSet<CellAddress> Dependents { get; set; } = new();
    public int Level { get; set; } = -1; // Topological level for sorting
    
    public DependencyNode(CellAddress address)
    {
        Address = address;
    }
}

/// <summary>
/// Manages the dependency graph (DAG) for efficient cell evaluation
/// </summary>
public class DependencyGraph
{
    private readonly Dictionary<CellAddress, DependencyNode> _nodes = new();
    private readonly Regex _cellReferenceRegex = new Regex(
        @"\b([A-Z]+[0-9]+)\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );
    private readonly Regex _rangeReferenceRegex = new Regex(
        @"\b([A-Z]+[0-9]+):([A-Z]+[0-9]+)\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

    /// <summary>
    /// Gets or creates a node for the given cell address
    /// </summary>
    private DependencyNode GetOrCreateNode(CellAddress address)
    {
        if (!_nodes.TryGetValue(address, out var node))
        {
            node = new DependencyNode(address);
            _nodes[address] = node;
        }
        return node;
    }

    /// <summary>
    /// Updates the dependencies for a cell based on its formula
    /// </summary>
    public void UpdateCellDependencies(CellAddress address, string? formula)
    {
        var node = GetOrCreateNode(address);
        
        // Clear old dependencies
        foreach (var dep in node.Dependencies.ToList())
        {
            if (_nodes.TryGetValue(dep, out var depNode))
            {
                depNode.Dependents.Remove(address);
            }
        }
        node.Dependencies.Clear();

        if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("="))
        {
            return;
        }

        // Extract cell references from formula
        var references = ExtractCellReferences(formula);
        foreach (var reference in references)
        {
            node.Dependencies.Add(reference);
            var depNode = GetOrCreateNode(reference);
            depNode.Dependents.Add(address);
        }
    }

    /// <summary>
    /// Extracts all cell references from a formula string
    /// </summary>
    private HashSet<CellAddress> ExtractCellReferences(string formula)
    {
        var references = new HashSet<CellAddress>();

        // Extract range references (e.g., A1:A10)
        var rangeMatches = _rangeReferenceRegex.Matches(formula);
        foreach (Match match in rangeMatches)
        {
            if (CellAddress.TryParse(match.Groups[1].Value, "Sheet1", out var start) &&
                CellAddress.TryParse(match.Groups[2].Value, "Sheet1", out var end))
            {
                // Add all cells in the range
                for (int row = start.Row; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        references.Add(new CellAddress(start.SheetName, row, col));
                    }
                }
            }
        }

        // Extract single cell references (e.g., A1, B2)
        var cellMatches = _cellReferenceRegex.Matches(formula);
        foreach (Match match in cellMatches)
        {
            // Skip if this was part of a range (already processed)
            if (rangeMatches.Cast<Match>().Any(r => r.Index <= match.Index && match.Index < r.Index + r.Length))
            {
                continue;
            }

            var cellRef = match.Groups[1].Value;
            if (CellAddress.TryParse(cellRef, "Sheet1", out var address))
            {
                references.Add(address);
            }
        }

        return references;
    }

    /// <summary>
    /// Removes a cell from the dependency graph
    /// </summary>
    public void RemoveCell(CellAddress address)
    {
        if (!_nodes.TryGetValue(address, out var node))
        {
            return;
        }

        // Remove from dependents of cells it depends on
        foreach (var dep in node.Dependencies)
        {
            if (_nodes.TryGetValue(dep, out var depNode))
            {
                depNode.Dependents.Remove(address);
            }
        }

        // Remove from dependencies of cells that depend on it
        foreach (var dependent in node.Dependents)
        {
            if (_nodes.TryGetValue(dependent, out var depNode))
            {
                depNode.Dependencies.Remove(address);
            }
        }

        _nodes.Remove(address);
    }

    /// <summary>
    /// Gets all cells that directly depend on the given cell
    /// </summary>
    public IEnumerable<CellAddress> GetDirectDependents(CellAddress address)
    {
        if (_nodes.TryGetValue(address, out var node))
        {
            return node.Dependents;
        }
        return Enumerable.Empty<CellAddress>();
    }

    /// <summary>
    /// Gets all cells that the given cell depends on (directly)
    /// </summary>
    public IEnumerable<CellAddress> GetDirectDependencies(CellAddress address)
    {
        if (_nodes.TryGetValue(address, out var node))
        {
            return node.Dependencies;
        }
        return Enumerable.Empty<CellAddress>();
    }

    /// <summary>
    /// Gets all cells that transitively depend on the given cell
    /// </summary>
    public HashSet<CellAddress> GetAllDependents(CellAddress address)
    {
        var result = new HashSet<CellAddress>();
        var queue = new Queue<CellAddress>();
        queue.Enqueue(address);

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            var directDependents = GetDirectDependents(current);
            
            foreach (var dependent in directDependents)
            {
                if (result.Add(dependent))
                {
                    queue.Enqueue(dependent);
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Detects circular references starting from the given cell
    /// </summary>
    public List<CellAddress>? DetectCircularReference(CellAddress startAddress)
    {
        var visited = new HashSet<CellAddress>();
        var recursionStack = new HashSet<CellAddress>();
        var path = new List<CellAddress>();

        if (HasCircularReference(startAddress, visited, recursionStack, path))
        {
            return path;
        }

        return null;
    }

    private bool HasCircularReference(
        CellAddress address,
        HashSet<CellAddress> visited,
        HashSet<CellAddress> recursionStack,
        List<CellAddress> path)
    {
        if (!visited.Contains(address))
        {
            visited.Add(address);
            recursionStack.Add(address);
            path.Add(address);

            var dependencies = GetDirectDependencies(address);
            foreach (var dep in dependencies)
            {
                if (!visited.Contains(dep))
                {
                    if (HasCircularReference(dep, visited, recursionStack, path))
                    {
                        return true;
                    }
                }
                else if (recursionStack.Contains(dep))
                {
                    // Found a cycle
                    path.Add(dep);
                    return true;
                }
            }
        }

        recursionStack.Remove(address);
        if (path.Count > 0 && path[path.Count - 1].Equals(address))
        {
            path.RemoveAt(path.Count - 1);
        }
        return false;
    }

    /// <summary>
    /// Performs topological sort to determine evaluation order
    /// Returns list of batches where each batch can be evaluated in parallel
    /// </summary>
    public List<List<CellAddress>> GetEvaluationOrder()
    {
        // Reset all levels
        foreach (var node in _nodes.Values)
        {
            node.Level = -1;
        }

        // Calculate levels using Kahn's algorithm
        var inDegree = new Dictionary<CellAddress, int>();
        foreach (var node in _nodes.Values)
        {
            inDegree[node.Address] = node.Dependencies.Count;
        }

        var queue = new Queue<DependencyNode>();
        foreach (var node in _nodes.Values.Where(n => inDegree[n.Address] == 0))
        {
            node.Level = 0;
            queue.Enqueue(node);
        }

        while (queue.Count > 0)
        {
            var node = queue.Dequeue();
            
            foreach (var dependent in node.Dependents)
            {
                if (_nodes.TryGetValue(dependent, out var depNode))
                {
                    inDegree[dependent]--;
                    
                    if (inDegree[dependent] == 0)
                    {
                        depNode.Level = node.Level + 1;
                        queue.Enqueue(depNode);
                    }
                }
            }
        }

        // Group nodes by level
        var batches = new List<List<CellAddress>>();
        
        // Handle empty graph
        if (_nodes.Count == 0)
        {
            return batches;
        }
        
        var maxLevel = _nodes.Values.Max(n => n.Level);
        
        for (int level = 0; level <= maxLevel; level++)
        {
            var batch = _nodes.Values
                .Where(n => n.Level == level)
                .Select(n => n.Address)
                .ToList();
            
            if (batch.Count > 0)
            {
                batches.Add(batch);
            }
        }

        return batches;
    }

    /// <summary>
    /// Gets all cells in the graph
    /// </summary>
    public IEnumerable<CellAddress> GetAllCells()
    {
        return _nodes.Keys;
    }

    /// <summary>
    /// Clears the entire dependency graph
    /// </summary>
    public void Clear()
    {
        _nodes.Clear();
    }

    /// <summary>
    /// Gets statistics about the dependency graph
    /// </summary>
    public (int NodeCount, int EdgeCount, int MaxDepth) GetStatistics()
    {
        var nodeCount = _nodes.Count;
        var edgeCount = _nodes.Values.Sum(n => n.Dependencies.Count);
        var maxDepth = _nodes.Values.Any() ? _nodes.Values.Max(n => n.Level) + 1 : 0;

        return (nodeCount, edgeCount, maxDepth);
    }
}
