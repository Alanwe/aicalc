using System;
using System.Collections.Generic;
using System.Linq;
using AiCalc.Models;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// Manages undo/redo operations for cell changes (Phase 5)
/// </summary>
public class UndoRedoManager
{
    private readonly Stack<CellChangeAction> _undoStack = new();
    private readonly Stack<CellChangeAction> _redoStack = new();
    private readonly int _maxUndoLevels;
    private bool _isUndoRedoInProgress;

    public event EventHandler? StacksChanged;

    public bool CanUndo => _undoStack.Count > 0;
    public bool CanRedo => _redoStack.Count > 0;
    public int UndoCount => _undoStack.Count;
    public int RedoCount => _redoStack.Count;

    public UndoRedoManager(int maxUndoLevels = 50)
    {
        _maxUndoLevels = maxUndoLevels;
    }

    /// <summary>
    /// Record a new action (clears redo stack)
    /// </summary>
    public void RecordAction(CellChangeAction action)
    {
        if (_isUndoRedoInProgress)
        {
            return; // Don't record actions during undo/redo
        }

        _undoStack.Push(action);
        _redoStack.Clear();

        // Limit stack size
        if (_undoStack.Count > _maxUndoLevels)
        {
            var temp = _undoStack.Reverse().Take(_maxUndoLevels).Reverse().ToList();
            _undoStack.Clear();
            foreach (var item in temp)
            {
                _undoStack.Push(item);
            }
        }

        StacksChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Undo the last action and return it
    /// </summary>
    public CellChangeAction? Undo()
    {
        if (!CanUndo)
        {
            return null;
        }

        _isUndoRedoInProgress = true;
        try
        {
            var action = _undoStack.Pop();
            _redoStack.Push(action);
            StacksChanged?.Invoke(this, EventArgs.Empty);
            return action;
        }
        finally
        {
            _isUndoRedoInProgress = false;
        }
    }

    /// <summary>
    /// Redo the last undone action and return it
    /// </summary>
    public CellChangeAction? Redo()
    {
        if (!CanRedo)
        {
            return null;
        }

        _isUndoRedoInProgress = true;
        try
        {
            var action = _redoStack.Pop();
            _undoStack.Push(action);
            StacksChanged?.Invoke(this, EventArgs.Empty);
            return action;
        }
        finally
        {
            _isUndoRedoInProgress = false;
        }
    }

    /// <summary>
    /// Clear all undo/redo history
    /// </summary>
    public void Clear()
    {
        _undoStack.Clear();
        _redoStack.Clear();
        StacksChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Get peek at the next undo action without removing it
    /// </summary>
    public CellChangeAction? PeekUndo()
    {
        return CanUndo ? _undoStack.Peek() : null;
    }

    /// <summary>
    /// Get peek at the next redo action without removing it
    /// </summary>
    public CellChangeAction? PeekRedo()
    {
        return CanRedo ? _redoStack.Peek() : null;
    }
}
