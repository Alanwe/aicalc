using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// Progress information for cell evaluation
/// </summary>
public class EvaluationProgress
{
    public int TotalCells { get; set; }
    public int CompletedCells { get; set; }
    public int CurrentBatch { get; set; }
    public int TotalBatches { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public double PercentComplete => TotalCells > 0 ? (double)CompletedCells / TotalCells * 100 : 0;
}

/// <summary>
/// Result of an evaluation operation
/// </summary>
public class EvaluationResult
{
    public bool Success { get; set; }
    public int CellsEvaluated { get; set; }
    public int CellsFailed { get; set; }
    public TimeSpan Duration { get; set; }
    public List<(CellAddress Address, string Error)> Errors { get; set; } = new();
}

/// <summary>
/// Multi-threaded cell evaluation engine with dependency management
/// </summary>
public class EvaluationEngine
{
    private readonly DependencyGraph _dependencyGraph;
    private readonly FunctionRunner _functionRunner;
    private int _maxDegreeOfParallelism;
    private int _defaultTimeoutSeconds;
    
    public event EventHandler<EvaluationProgress>? ProgressChanged;
    public event EventHandler<(CellAddress Address, string? OldValue, string? NewValue)>? CellValueChanged;

    /// <summary>
    /// Maximum number of cells to evaluate in parallel. Configurable from Settings menu (Task 9).
    /// </summary>
    public int MaxDegreeOfParallelism
    {
        get => _maxDegreeOfParallelism;
        set => _maxDegreeOfParallelism = value > 0 ? value : Environment.ProcessorCount;
    }

    /// <summary>
    /// Default timeout in seconds for cell evaluation. Default 100 seconds per Task 9.
    /// Configurable at AI Service definition.
    /// </summary>
    public int DefaultTimeoutSeconds
    {
        get => _defaultTimeoutSeconds;
        set => _defaultTimeoutSeconds = value > 0 ? value : 100;
    }

    public EvaluationEngine(
        DependencyGraph dependencyGraph,
        FunctionRunner functionRunner,
        int? maxDegreeOfParallelism = null,
        int defaultTimeoutSeconds = 100) // Changed from 30 to 100 per Task 9 requirement
    {
        _dependencyGraph = dependencyGraph;
        _functionRunner = functionRunner;
        _maxDegreeOfParallelism = maxDegreeOfParallelism ?? Environment.ProcessorCount;
        _defaultTimeoutSeconds = defaultTimeoutSeconds;
    }

    /// <summary>
    /// Evaluates a single cell
    /// </summary>
    public async Task<bool> EvaluateCellAsync(
        CellViewModel cell,
        CancellationToken cancellationToken = default)
    {
        try
        {
            // Check for circular reference
            var circularRef = _dependencyGraph.DetectCircularReference(cell.Address);
            if (circularRef != null)
            {
                var cycle = string.Join(" → ", circularRef.Select(a => a.ToString()));
                cell.Value = new CellValue(CellObjectType.Error, $"#CIRCULAR! {cycle}", $"#CIRCULAR! {cycle}");
                return false;
            }

            // If not a formula, nothing to evaluate
            if (string.IsNullOrWhiteSpace(cell.Formula) || !cell.Formula.StartsWith("="))
            {
                return true;
            }

            // Store old value for change detection
            var oldValue = cell.DisplayValue;

            // Execute the formula
            var result = await _functionRunner.EvaluateAsync(cell, cell.Formula);

            if (result != null)
            {
                cell.Value = result.Value;
                
                // Notify if value changed
                var newValue = cell.DisplayValue;
                if (oldValue != newValue)
                {
                    CellValueChanged?.Invoke(this, (cell.Address, oldValue, newValue));
                }
                
                return true;
            }
            else
            {
                cell.Value = new CellValue(CellObjectType.Error, "#ERROR!", "#ERROR!");
                return false;
            }
        }
        catch (OperationCanceledException)
        {
            cell.Value = new CellValue(CellObjectType.Error, "#TIMEOUT!", "#TIMEOUT! Operation exceeded time limit");
            return false;
        }
        catch (Exception ex)
        {
            cell.Value = new CellValue(CellObjectType.Error, $"#ERROR! {ex.Message}", $"#ERROR! {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Evaluates all cells in the workbook in dependency order
    /// </summary>
    public async Task<EvaluationResult> EvaluateAllAsync(
        Dictionary<CellAddress, CellViewModel> cells,
        CancellationToken cancellationToken = default,
        IProgress<EvaluationProgress>? progress = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new EvaluationResult();
        
        // Get evaluation order (batches that can be parallelized)
        var batches = _dependencyGraph.GetEvaluationOrder();
        var totalCells = batches.Sum(b => b.Count);
        var completedCells = 0;

        for (int batchIndex = 0; batchIndex < batches.Count; batchIndex++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }

            var batch = batches[batchIndex];
            
            // Report progress
            var progressInfo = new EvaluationProgress
            {
                TotalCells = totalCells,
                CompletedCells = completedCells,
                CurrentBatch = batchIndex + 1,
                TotalBatches = batches.Count,
                ElapsedTime = stopwatch.Elapsed
            };
            progress?.Report(progressInfo);
            ProgressChanged?.Invoke(this, progressInfo);

            // Evaluate all cells in this batch in parallel
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = _maxDegreeOfParallelism,
                CancellationToken = cancellationToken
            };

            var batchTasks = batch
                .Where(address => cells.ContainsKey(address))
                .Select(async address =>
                {
                    var cell = cells[address];
                    var success = await EvaluateCellAsync(cell, cancellationToken);
                    var errorMsg = cell.Value.ObjectType == CellObjectType.Error ? cell.Value.DisplayValue : null;
                    return (address, success, errorMsg);
                })
                .ToList();

            var batchResults = await Task.WhenAll(batchTasks);

            foreach (var (address, success, error) in batchResults)
            {
                completedCells++;
                
                if (success)
                {
                    result.CellsEvaluated++;
                }
                else
                {
                    result.CellsFailed++;
                    if (error != null)
                    {
                        result.Errors.Add((address, error));
                    }
                }
            }
        }

        stopwatch.Stop();
        result.Duration = stopwatch.Elapsed;
        result.Success = result.CellsFailed == 0;

        // Final progress report
        var finalProgress = new EvaluationProgress
        {
            TotalCells = totalCells,
            CompletedCells = completedCells,
            CurrentBatch = batches.Count,
            TotalBatches = batches.Count,
            ElapsedTime = stopwatch.Elapsed
        };
        progress?.Report(finalProgress);
        ProgressChanged?.Invoke(this, finalProgress);

        return result;
    }

    /// <summary>
    /// Evaluates only the cells that depend on the given cell (cascade evaluation)
    /// </summary>
    public async Task<EvaluationResult> EvaluateDependentsAsync(
        CellAddress changedCell,
        Dictionary<CellAddress, CellViewModel> cells,
        CancellationToken cancellationToken = default,
        IProgress<EvaluationProgress>? progress = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new EvaluationResult();

        // Get all cells that depend on this cell
        var dependents = _dependencyGraph.GetAllDependents(changedCell);
        
        if (dependents.Count == 0)
        {
            result.Success = true;
            result.Duration = stopwatch.Elapsed;
            return result;
        }

        // Create a subgraph for just these cells
        var subgraphBatches = GetSubgraphEvaluationOrder(dependents);
        var totalCells = subgraphBatches.Sum(b => b.Count);
        var completedCells = 0;

        for (int batchIndex = 0; batchIndex < subgraphBatches.Count; batchIndex++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }

            var batch = subgraphBatches[batchIndex];
            
            // Report progress
            var progressInfo = new EvaluationProgress
            {
                TotalCells = totalCells,
                CompletedCells = completedCells,
                CurrentBatch = batchIndex + 1,
                TotalBatches = subgraphBatches.Count,
                ElapsedTime = stopwatch.Elapsed
            };
            progress?.Report(progressInfo);
            ProgressChanged?.Invoke(this, progressInfo);

            // Evaluate batch in parallel
            var batchTasks = batch
                .Where(address => cells.ContainsKey(address))
                .Select(async address =>
                {
                    var cell = cells[address];
                    var success = await EvaluateCellAsync(cell, cancellationToken);
                    var errorMsg = cell.Value.ObjectType == CellObjectType.Error ? cell.Value.DisplayValue : null;
                    return (address, success, errorMsg);
                })
                .ToList();

            var batchResults = await Task.WhenAll(batchTasks);

            foreach (var (address, success, error) in batchResults)
            {
                completedCells++;
                
                if (success)
                {
                    result.CellsEvaluated++;
                }
                else
                {
                    result.CellsFailed++;
                    if (error != null)
                    {
                        result.Errors.Add((address, error));
                    }
                }
            }
        }

        stopwatch.Stop();
        result.Duration = stopwatch.Elapsed;
        result.Success = result.CellsFailed == 0;

        return result;
    }

    /// <summary>
    /// Gets evaluation order for a subset of cells
    /// </summary>
    private List<List<CellAddress>> GetSubgraphEvaluationOrder(HashSet<CellAddress> cells)
    {
        // Build a temporary level map for just these cells
        var levels = new Dictionary<CellAddress, int>();
        var processed = new HashSet<CellAddress>();
        var queue = new Queue<CellAddress>();

        // Find cells with no dependencies (or dependencies outside the subgraph)
        foreach (var cell in cells)
        {
            var deps = _dependencyGraph.GetDirectDependencies(cell);
            var depsInSubgraph = deps.Count(d => cells.Contains(d));
            
            if (depsInSubgraph == 0)
            {
                levels[cell] = 0;
                queue.Enqueue(cell);
                processed.Add(cell);
            }
        }

        // Process remaining cells
        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            var currentLevel = levels[current];

            var dependents = _dependencyGraph.GetDirectDependents(current)
                .Where(d => cells.Contains(d));

            foreach (var dependent in dependents)
            {
                if (processed.Contains(dependent))
                {
                    continue;
                }

                var deps = _dependencyGraph.GetDirectDependencies(dependent)
                    .Where(d => cells.Contains(d));

                if (deps.All(d => processed.Contains(d)))
                {
                    var maxDepLevel = deps.Select(d => levels[d]).Max();
                    levels[dependent] = maxDepLevel + 1;
                    queue.Enqueue(dependent);
                    processed.Add(dependent);
                }
            }
        }

        // Group by level
        var batches = new List<List<CellAddress>>();
        if (levels.Any())
        {
            var maxLevel = levels.Values.Max();
            for (int level = 0; level <= maxLevel; level++)
            {
                var batch = levels.Where(kvp => kvp.Value == level)
                    .Select(kvp => kvp.Key)
                    .ToList();
                
                if (batch.Count > 0)
                {
                    batches.Add(batch);
                }
            }
        }

        return batches;
    }

    /// <summary>
    /// Cancels any ongoing evaluation
    /// </summary>
    public void CancelEvaluation(CancellationTokenSource cts)
    {
        cts?.Cancel();
    }
}
