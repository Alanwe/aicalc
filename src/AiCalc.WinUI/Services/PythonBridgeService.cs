using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// IPC bridge service for Python SDK communication (Phase 7)
/// Uses Named Pipes for secure inter-process communication
/// </summary>
public class PythonBridgeService : IDisposable
{
    private readonly WorkbookViewModel _workbook;
    private NamedPipeServerStream? _pipeServer;
    private CancellationTokenSource? _cancellationTokenSource;
    private Task? _serverTask;
    private readonly string _pipeName;
    private bool _disposed;

    public event EventHandler<string>? MessageReceived;
    public event EventHandler<Exception>? ErrorOccurred;

    public bool IsRunning { get; private set; }

    public PythonBridgeService(WorkbookViewModel workbook, string pipeName = "AiCalc_Bridge")
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _pipeName = pipeName;
    }

    /// <summary>
    /// Start the IPC server
    /// </summary>
    public void Start()
    {
        var logPath = Path.Combine(Path.GetTempPath(), "aicalc_python_bridge.log");
        try
        {
            File.AppendAllText(logPath, $"\n=== START CALLED at {DateTime.Now} ===\n");
        }
        catch
        {
            // Ignore file write errors
        }

        if (IsRunning)
        {
            try { File.AppendAllText(logPath, "Already running, returning\n"); } catch { }
            return;
        }

        try { File.AppendAllText(logPath, $"\n[{DateTime.Now}] Start() called\n"); } catch { }

        _cancellationTokenSource = new CancellationTokenSource();
        IsRunning = true;
        
        // Start server loop in background
        _serverTask = Task.Run(async () =>
        {
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] Task.Run started\n");
                System.Diagnostics.Debug.WriteLine($"[PythonBridge] Starting server on pipe: {_pipeName}");
                MessageReceived?.Invoke(this, $"Python bridge server starting on pipe: {_pipeName}");
                await RunServerAsync(_cancellationTokenSource.Token);
                File.AppendAllText(logPath, $"[{DateTime.Now}] RunServerAsync completed\n");
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] FATAL ERROR: {ex}\n");
                System.Diagnostics.Debug.WriteLine($"[PythonBridge] FATAL ERROR: {ex}");
                ErrorOccurred?.Invoke(this, new Exception($"Server startup failed: {ex.Message}", ex));
            }
        }, _cancellationTokenSource.Token);
        
        try { File.AppendAllText(logPath, $"[{DateTime.Now}] Start() completed, log: {logPath}\n"); } catch { }
        System.Diagnostics.Debug.WriteLine($"[PythonBridge] Service started, log: {logPath}");
        MessageReceived?.Invoke(this, $"Python bridge service started, log: {logPath}");
    }

    /// <summary>
    /// Stop the IPC server
    /// </summary>
    public void Stop()
    {
        if (!IsRunning) return;

        MessageReceived?.Invoke(this, "Python bridge service stopping");
        _cancellationTokenSource?.Cancel();
        _pipeServer?.Dispose();
        IsRunning = false;
    }

    private async Task RunServerAsync(CancellationToken cancellationToken)
    {
        var logPath = Path.Combine(Path.GetTempPath(), "aicalc_python_bridge.log");
        File.AppendAllText(logPath, $"[{DateTime.Now}] RunServerAsync started\n");
        MessageReceived?.Invoke(this, "Python bridge server loop started");
        
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] Creating pipe server instance...\n");
                MessageReceived?.Invoke(this, "Creating new pipe server instance...");
                _pipeServer = new NamedPipeServerStream(
                    _pipeName,
                    PipeDirection.InOut,
                    NamedPipeServerStream.MaxAllowedServerInstances,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                File.AppendAllText(logPath, $"[{DateTime.Now}] Pipe created, waiting for connection...\n");
                MessageReceived?.Invoke(this, $"Waiting for Python client connection on {_pipeName}...");
                
                // Wait for client connection
                await _pipeServer.WaitForConnectionAsync(cancellationToken);
                
                File.AppendAllText(logPath, $"[{DateTime.Now}] Client connected!\n");
                MessageReceived?.Invoke(this, "Python client connected!");

                // Handle client requests
                await HandleClientAsync(_pipeServer, cancellationToken);
                
                File.AppendAllText(logPath, $"[{DateTime.Now}] Client disconnected\n");
                MessageReceived?.Invoke(this, "Python client disconnected");
            }
            catch (OperationCanceledException)
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] Operation cancelled\n");
                MessageReceived?.Invoke(this, "Server operation cancelled");
                break;
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] Server error: {ex}\n");
                ErrorOccurred?.Invoke(this, new Exception($"Server error: {ex.Message}", ex));
                // Wait a bit before retrying
                await Task.Delay(1000, cancellationToken);
            }
            finally
            {
                _pipeServer?.Dispose();
                _pipeServer = null;
            }
        }
        
        File.AppendAllText(logPath, $"[{DateTime.Now}] RunServerAsync loop ended\n");
        MessageReceived?.Invoke(this, "Python bridge server loop ended");
    }

    private async Task HandleClientAsync(NamedPipeServerStream pipe, CancellationToken cancellationToken)
    {
        var logPath = Path.Combine(Path.GetTempPath(), "aicalc_python_bridge.log");
        
        try { File.AppendAllText(logPath, $"[{DateTime.Now}] HandleClientAsync started\n"); } catch { }

        var buffer = new byte[4096];
        var messageBuilder = new StringBuilder();

        while (pipe.IsConnected && !cancellationToken.IsCancellationRequested)
        {
            try
            {
                try { File.AppendAllText(logPath, $"[{DateTime.Now}] Reading from pipe...\n"); } catch { }
                
                // Read until we get a newline
                int bytesRead = await pipe.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                if (bytesRead == 0) break;
                
                var chunk = Encoding.UTF8.GetString(buffer, 0, bytesRead);
                messageBuilder.Append(chunk);
                
                try { File.AppendAllText(logPath, $"[{DateTime.Now}] Read {bytesRead} bytes: {chunk}\n"); } catch { }
                
                // Check if we have a complete message (ends with \n)
                var message = messageBuilder.ToString();
                if (!message.Contains('\n')) continue;
                
                // Process complete messages
                var lines = message.Split('\n');
                for (int i = 0; i < lines.Length - 1; i++)
                {
                    var request = lines[i].Trim();
                    if (string.IsNullOrEmpty(request)) continue;

                    try { File.AppendAllText(logPath, $"[{DateTime.Now}] Processing: {request}\n"); } catch { }
                    MessageReceived?.Invoke(this, request);

                    var response = await ProcessRequestAsync(request);
                    try { File.AppendAllText(logPath, $"[{DateTime.Now}] Response: {response}\n"); } catch { }
                    
                    var responseBytes = Encoding.UTF8.GetBytes(response + "\n");
                    await pipe.WriteAsync(responseBytes, 0, responseBytes.Length, cancellationToken);
                    await pipe.FlushAsync(cancellationToken);
                    
                    try { File.AppendAllText(logPath, $"[{DateTime.Now}] Response sent\n"); } catch { }
                }
                
                // Keep the incomplete part
                messageBuilder.Clear();
                if (lines.Length > 0)
                {
                    messageBuilder.Append(lines[lines.Length - 1]);
                }
            }
            catch (IOException ex)
            {
                try { File.AppendAllText(logPath, $"[{DateTime.Now}] IOException: {ex.Message}\n"); } catch { }
                break;
            }
            catch (Exception ex)
            {
                try { File.AppendAllText(logPath, $"[{DateTime.Now}] Exception: {ex}\n"); } catch { }
                ErrorOccurred?.Invoke(this, ex);
                
                var errorResponse = JsonSerializer.Serialize(new { success = false, error = ex.Message });
                var errorBytes = Encoding.UTF8.GetBytes(errorResponse + "\n");
                await pipe.WriteAsync(errorBytes, 0, errorBytes.Length, cancellationToken);
            }
        }
        
        try { File.AppendAllText(logPath, $"[{DateTime.Now}] HandleClientAsync ended\n"); } catch { }
    }

    private async Task<string> ProcessRequestAsync(string requestJson)
    {
        try
        {
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var request = JsonSerializer.Deserialize<PythonRequest>(requestJson, options);
            if (request == null)
            {
                return CreateErrorResponse("Invalid request format");
            }

            return request.Command switch
            {
                "get_value" => await GetValueAsync(request),
                "set_value" => await SetValueAsync(request),
                "get_range" => await GetRangeAsync(request),
                "run_function" => await RunFunctionAsync(request),
                "get_sheets" => GetSheets(),
                "ping" => CreateSuccessResponse("pong"),
                _ => CreateErrorResponse($"Unknown command: {request.Command}")
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex.Message);
        }
    }

    private Task<string> GetValueAsync(PythonRequest request)
    {
        if (string.IsNullOrEmpty(request.CellRef))
        {
            return Task.FromResult(CreateErrorResponse("CellRef is required"));
        }

        var cell = FindCell(request.CellRef);
        if (cell == null)
        {
            return Task.FromResult(CreateErrorResponse($"Cell not found: {request.CellRef}"));
        }

        return Task.FromResult(CreateSuccessResponse(new
        {
            cell_ref = request.CellRef,
            value = cell.Value.DisplayValue,
            serialized_value = cell.Value.SerializedValue,
            object_type = cell.Value.ObjectType.ToString(),
            formula = cell.Formula
        }));
    }

    private Task<string> SetValueAsync(PythonRequest request)
    {
        if (string.IsNullOrEmpty(request.CellRef))
        {
            return Task.FromResult(CreateErrorResponse("CellRef is required"));
        }

        var cell = FindCell(request.CellRef);
        if (cell == null)
        {
            return Task.FromResult(CreateErrorResponse($"Cell not found: {request.CellRef}"));
        }

        cell.RawValue = request.Value?.ToString();
        return Task.FromResult(CreateSuccessResponse(new { cell_ref = request.CellRef }));
    }

    private Task<string> GetRangeAsync(PythonRequest request)
    {
        if (string.IsNullOrEmpty(request.RangeRef))
        {
            return Task.FromResult(CreateErrorResponse("RangeRef is required"));
        }

        // TODO: Implement range parsing and data extraction
        return Task.FromResult(CreateErrorResponse("Range operations not yet implemented"));
    }

    private async Task<string> RunFunctionAsync(PythonRequest request)
    {
        if (string.IsNullOrEmpty(request.FunctionName))
        {
            return CreateErrorResponse("FunctionName is required");
        }

        try
        {
            if (!_workbook.FunctionRegistry.TryGet(request.FunctionName, out var descriptor))
            {
                return CreateErrorResponse($"Function not found: {request.FunctionName}");
            }

            var sheet = _workbook.Sheets.FirstOrDefault();
            if (sheet == null)
            {
                return CreateErrorResponse("No sheets available in workbook");
            }

            // Convert args to cell view models (empty cells for literals)
            var argCells = new List<CellViewModel>();
            // For now, we'll pass empty arguments - proper implementation would need cell resolution
            
            var formula = $"={request.FunctionName}({string.Join(",", request.Args ?? Array.Empty<object>())})";
            var context = new FunctionEvaluationContext(
                _workbook,
                sheet,
                argCells,
                formula);

            var result = await descriptor.Handler(context);

            return CreateSuccessResponse(new
            {
                function_name = request.FunctionName,
                result = result.Value.DisplayValue,
                serialized_value = result.Value.SerializedValue,
                object_type = result.Value.ObjectType.ToString()
            });
        }
        catch (Exception ex)
        {
            return CreateErrorResponse($"Function execution failed: {ex.Message}");
        }
    }

    private string GetSheets()
    {
        var sheets = _workbook.Sheets.Select(s => new
        {
            name = s.Name,
            row_count = s.Rows.Count,
            column_count = s.ColumnCount
        }).ToArray();

        return CreateSuccessResponse(new { sheets });
    }

    private CellViewModel? FindCell(string cellRef)
    {
        try
        {
            var defaultSheet = _workbook.SelectedSheet?.Name ?? _workbook.Sheets.FirstOrDefault()?.Name ?? "Sheet1";
            if (!Models.CellAddress.TryParse(cellRef, defaultSheet, out var address))
            {
                return null;
            }
            
            var sheet = _workbook.Sheets.FirstOrDefault(s => s.Name == address.SheetName) 
                        ?? _workbook.Sheets.FirstOrDefault();
            
            return sheet?.GetCell(address.Row, address.Column);
        }
        catch
        {
            return null;
        }
    }

    private static string CreateSuccessResponse(object? data = null)
    {
        return JsonSerializer.Serialize(new
        {
            success = true,
            data
        });
    }

    private static string CreateErrorResponse(string error)
    {
        return JsonSerializer.Serialize(new
        {
            success = false,
            error
        });
    }

    public void Dispose()
    {
        if (_disposed) return;

        Stop();
        _cancellationTokenSource?.Dispose();
        _pipeServer?.Dispose();
        _disposed = true;
    }
}

/// <summary>
/// Request format from Python client
/// </summary>
public class PythonRequest
{
    public string Command { get; set; } = string.Empty;
    public string? CellRef { get; set; }
    public string? RangeRef { get; set; }
    public object? Value { get; set; }
    public string? FunctionName { get; set; }
    public object[]? Args { get; set; }
}
