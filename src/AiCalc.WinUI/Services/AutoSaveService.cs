using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// Provides automatic workbook saving functionality (Phase 6)
/// </summary>
public class AutoSaveService : IDisposable
{
    private readonly WorkbookViewModel _workbook;
    private Timer? _autoSaveTimer;
    private bool _isDirty;
    private bool _isEnabled;
    private int _intervalMinutes = 5;
    private string? _lastSavePath;
    private bool _disposed;

    public event EventHandler<string>? AutoSaved;
    public event EventHandler<Exception>? AutoSaveFailed;

    public bool IsEnabled
    {
        get => _isEnabled;
        set
        {
            if (_isEnabled != value)
            {
                _isEnabled = value;
                if (_isEnabled)
                {
                    StartTimer();
                }
                else
                {
                    StopTimer();
                }
            }
        }
    }

    public int IntervalMinutes
    {
        get => _intervalMinutes;
        set
        {
            if (value < 1) value = 1;
            if (value > 60) value = 60;
            
            if (_intervalMinutes != value)
            {
                _intervalMinutes = value;
                if (_isEnabled)
                {
                    StartTimer(); // Restart with new interval
                }
            }
        }
    }

    public AutoSaveService(WorkbookViewModel workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    public void MarkDirty()
    {
        _isDirty = true;
    }

    public void SetSavePath(string path)
    {
        _lastSavePath = path;
    }

    private void StartTimer()
    {
        StopTimer();
        _autoSaveTimer = new Timer(
            AutoSaveCallback,
            null,
            TimeSpan.FromMinutes(_intervalMinutes),
            TimeSpan.FromMinutes(_intervalMinutes));
    }

    private void StopTimer()
    {
        _autoSaveTimer?.Dispose();
        _autoSaveTimer = null;
    }

    private async void AutoSaveCallback(object? state)
    {
        if (!_isDirty || string.IsNullOrEmpty(_lastSavePath))
        {
            return;
        }

        try
        {
            // Create autosave backup path
            var directory = Path.GetDirectoryName(_lastSavePath);
            var fileName = Path.GetFileNameWithoutExtension(_lastSavePath);
            var extension = Path.GetExtension(_lastSavePath);
            var autoSavePath = Path.Combine(directory ?? ".", $"{fileName}_autosave{extension}");

            // Save workbook
            var definition = _workbook.ToDefinition();
            var json = System.Text.Json.JsonSerializer.Serialize(definition, new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase
            });

            await File.WriteAllTextAsync(autoSavePath, json);
            
            _isDirty = false;
            AutoSaved?.Invoke(this, autoSavePath);
        }
        catch (Exception ex)
        {
            AutoSaveFailed?.Invoke(this, ex);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        
        StopTimer();
        _disposed = true;
    }
}
