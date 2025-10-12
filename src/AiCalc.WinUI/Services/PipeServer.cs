using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using AiCalc.Models;
using AiCalc.ViewModels;
using Microsoft.UI.Dispatching;

namespace AiCalc.Services
{
    /// <summary>
    /// Named Pipe Server for Python SDK communication
    /// Handles IPC requests from Python clients
    /// </summary>
    public class PipeServer : IDisposable
    {
        private readonly string _pipeName;
        private readonly WorkbookViewModel _workbook;
        private readonly FunctionRunner _functionRunner;
        private readonly DispatcherQueue _dispatcherQueue;
        private CancellationTokenSource? _cancellationTokenSource;
        private Task? _serverTask;
        private bool _isRunning;

        public PipeServer(string pipeName, WorkbookViewModel workbook, FunctionRunner functionRunner, DispatcherQueue dispatcherQueue)
        {
            _pipeName = pipeName;
            _workbook = workbook;
            _functionRunner = functionRunner;
            _dispatcherQueue = dispatcherQueue;
        }

        /// <summary>
        /// Start the pipe server
        /// </summary>
        public void Start()
        {
            if (_isRunning)
                return;

            _isRunning = true;
            _cancellationTokenSource = new CancellationTokenSource();
            _serverTask = Task.Run(() => ServerLoop(_cancellationTokenSource.Token));
        }

        /// <summary>
        /// Stop the pipe server
        /// </summary>
        public void Stop()
        {
            if (!_isRunning)
                return;

            _isRunning = false;
            _cancellationTokenSource?.Cancel();
            _serverTask?.Wait(TimeSpan.FromSeconds(2));
        }

        /// <summary>
        /// Main server loop - accepts client connections
        /// </summary>
        private async Task ServerLoop(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    using var pipeServer = new NamedPipeServerStream(
                        _pipeName,
                        PipeDirection.InOut,
                        NamedPipeServerStream.MaxAllowedServerInstances,
                        PipeTransmissionMode.Message,
                        PipeOptions.Asynchronous);

                    // Wait for client connection
                    await pipeServer.WaitForConnectionAsync(cancellationToken);

                    // Handle client in separate task (allows multiple clients)
                    _ = Task.Run(() => HandleClient(pipeServer), CancellationToken.None);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Pipe server error: {ex.Message}");
                    await Task.Delay(1000, cancellationToken);
                }
            }
        }

        /// <summary>
        /// Handle individual client connection
        /// </summary>
        private async Task HandleClient(NamedPipeServerStream pipeServer)
        {
            try
            {
                while (pipeServer.IsConnected)
                {
                    var message = await ReceiveMessage(pipeServer);
                    if (message == null)
                        break;

                    var response = await ProcessMessage(message);
                    await SendMessage(pipeServer, response);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Client handler error: {ex.Message}");
            }
        }

        /// <summary>
        /// Receive message from pipe
        /// </summary>
        private async Task<IPCMessage?> ReceiveMessage(NamedPipeServerStream pipe)
        {
            try
            {
                // Read 4-byte length header
                var lengthBuffer = new byte[4];
                var bytesRead = await pipe.ReadAsync(lengthBuffer, 0, 4);
                if (bytesRead != 4)
                    return null;

                var length = BitConverter.ToInt32(lengthBuffer, 0);
                if (length <= 0 || length > 1024 * 1024) // Max 1MB
                    return null;

                // Read message body
                var messageBuffer = new byte[length];
                bytesRead = await pipe.ReadAsync(messageBuffer, 0, length);
                if (bytesRead != length)
                    return null;

                var json = Encoding.UTF8.GetString(messageBuffer);
                return JsonSerializer.Deserialize<IPCMessage>(json);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Send message to pipe
        /// </summary>
        private async Task SendMessage(NamedPipeServerStream pipe, IPCMessage message)
        {
            try
            {
                var json = JsonSerializer.Serialize(message);
                var messageBytes = Encoding.UTF8.GetBytes(json);
                var lengthBytes = BitConverter.GetBytes(messageBytes.Length);

                // Send length header + message
                await pipe.WriteAsync(lengthBytes, 0, 4);
                await pipe.WriteAsync(messageBytes, 0, messageBytes.Length);
                await pipe.FlushAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Send message error: {ex.Message}");
            }
        }

        /// <summary>
        /// Process IPC message and generate response
        /// </summary>
        private async Task<IPCMessage> ProcessMessage(IPCMessage message)
        {
            try
            {
                var command = message.Command;
                var parameters = message.Parameters;

                object? result = command switch
                {
                    "GetValue" => await GetValue(parameters),
                    "SetValue" => await SetValue(parameters),
                    "GetFormula" => await GetFormula(parameters),
                    "SetFormula" => await SetFormula(parameters),
                    "GetRange" => await GetRange(parameters),
                    "RunFunction" => await RunFunction(parameters),
                    "EvaluateCell" => await EvaluateCell(parameters),
                    _ => new { status = "error", error = $"Unknown command: {command}" }
                };

                return new IPCMessage
                {
                    Command = message.Command + "Response",
                    Parameters = result as Dictionary<string, object> ?? new Dictionary<string, object> { ["result"] = result ?? "" },
                    RequestId = message.RequestId
                };
            }
            catch (Exception ex)
            {
                return new IPCMessage
                {
                    Command = message.Command + "Response",
                    Parameters = new Dictionary<string, object>
                    {
                        ["status"] = "error",
                        ["error"] = ex.Message
                    },
                    RequestId = message.RequestId
                };
            }
        }

        /// <summary>
        /// Get cell value
        /// </summary>
        private async Task<object> GetValue(Dictionary<string, object> parameters)
        {
            var sheetName = parameters["sheet"].ToString() ?? "Sheet1";
            var row = Convert.ToInt32(parameters["row"]);
            var column = Convert.ToInt32(parameters["column"]);

            object? value = null;
            var tcs = new TaskCompletionSource<bool>();
            
            _dispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    var sheet = _workbook.GetSheet(sheetName) ?? _workbook.Sheets.FirstOrDefault();
                    var cell = sheet?.GetCell(row, column);
                    value = cell?.Value?.DisplayValue ?? "";
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            
            await tcs.Task;
            return new Dictionary<string, object>
            {
                ["value"] = value ?? "",
                ["status"] = "success"
            };
        }

        /// <summary>
        /// Set cell value
        /// </summary>
        private async Task<object> SetValue(Dictionary<string, object> parameters)
        {
            var sheetName = parameters["sheet"].ToString() ?? "Sheet1";
            var row = Convert.ToInt32(parameters["row"]);
            var column = Convert.ToInt32(parameters["column"]);
            var value = parameters["value"];

            var tcs = new TaskCompletionSource<bool>();
            
            _dispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    var sheet = _workbook.GetSheet(sheetName) ?? _workbook.Sheets.FirstOrDefault();
                    var cell = sheet?.GetCell(row, column);
                    if (cell != null)
                    {
                        cell.RawValue = value?.ToString() ?? "";
                    }
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            
            await tcs.Task;
            return new Dictionary<string, object> { ["status"] = "success" };
        }

        /// <summary>
        /// Get cell formula
        /// </summary>
        private async Task<object> GetFormula(Dictionary<string, object> parameters)
        {
            var sheetName = parameters["sheet"].ToString() ?? "Sheet1";
            var row = Convert.ToInt32(parameters["row"]);
            var column = Convert.ToInt32(parameters["column"]);

            string? formula = null;
            var tcs = new TaskCompletionSource<bool>();
            
            _dispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    var sheet = _workbook.GetSheet(sheetName) ?? _workbook.Sheets.FirstOrDefault();
                    var cell = sheet?.GetCell(row, column);
                    formula = cell?.Formula;
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            
            await tcs.Task;
            return new Dictionary<string, object>
            {
                ["formula"] = formula ?? "",
                ["status"] = "success"
            };
        }

        /// <summary>
        /// Set cell formula
        /// </summary>
        private async Task<object> SetFormula(Dictionary<string, object> parameters)
        {
            var sheetName = parameters["sheet"].ToString() ?? "Sheet1";
            var row = Convert.ToInt32(parameters["row"]);
            var column = Convert.ToInt32(parameters["column"]);
            var formula = parameters["formula"].ToString() ?? "";

            var tcs = new TaskCompletionSource<bool>();
            
            _dispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    var sheet = _workbook.GetSheet(sheetName) ?? _workbook.Sheets.FirstOrDefault();
                    var cell = sheet?.GetCell(row, column);
                    if (cell != null)
                    {
                        cell.RawValue = formula;
                    }
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            
            await tcs.Task;
            return new Dictionary<string, object> { ["status"] = "success" };
        }

        /// <summary>
        /// Get range of cells
        /// </summary>
        private async Task<object> GetRange(Dictionary<string, object> parameters)
        {
            var rangeRef = parameters["range"].ToString() ?? "A1";
            var rangeParts = rangeRef.Split(':');
            
            if (!CellAddress.TryParse(rangeParts[0], "Sheet1", out var start))
            {
                throw new ArgumentException($"Invalid cell reference: {rangeParts[0]}");
            }
            
            var end = rangeParts.Length > 1 && CellAddress.TryParse(rangeParts[1], start.SheetName, out var endAddr)
                ? endAddr
                : start;

            var values = new List<List<object>>();
            var tcs = new TaskCompletionSource<bool>();

            _dispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    var sheet = _workbook.GetSheet(start.SheetName) ?? _workbook.Sheets.FirstOrDefault();
                    if (sheet != null)
                    {
                        for (int row = start.Row; row <= end.Row; row++)
                        {
                            var rowValues = new List<object>();
                            for (int col = start.Column; col <= end.Column; col++)
                            {
                                var cell = sheet.GetCell(row, col);
                                rowValues.Add(cell?.Value?.DisplayValue ?? "");
                            }
                            values.Add(rowValues);
                        }
                    }
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            
            await tcs.Task;
            return new Dictionary<string, object>
            {
                ["values"] = values,
                ["status"] = "success"
            };
        }

        /// <summary>
        /// Run AiCalc function
        /// </summary>
        private async Task<object> RunFunction(Dictionary<string, object> parameters)
        {
            var functionName = parameters["function"].ToString() ?? "";
            var arguments = parameters.ContainsKey("arguments") ? 
                (parameters["arguments"] as IEnumerable<object> ?? Array.Empty<object>()).Select(a => a?.ToString() ?? "").ToArray() : 
                Array.Empty<string>();

            string? resultValue = null;
            Exception? error = null;
            var tcs = new TaskCompletionSource<bool>();
            
            _dispatcherQueue.TryEnqueue(async () =>
            {
                try
                {
                    var sheet = _workbook.Sheets.FirstOrDefault();
                    if (sheet != null)
                    {
                        var argStr = string.Join(", ", arguments.Select(a => $"\"{a}\""));
                        var formula = $"={functionName}({argStr})";
                        
                        var tempCell = sheet.GetCell(0, 0);
                        if (tempCell != null)
                        {
                            var funcResult = await _functionRunner.EvaluateAsync(tempCell, formula);
                            resultValue = funcResult?.Value?.DisplayValue ?? "";
                        }
                    }
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    error = ex;
                    tcs.SetResult(true);
                }
            });
            
            await tcs.Task;

            if (error != null)
            {
                return new Dictionary<string, object>
                {
                    ["status"] = "error",
                    ["error"] = error.Message
                };
            }

            return new Dictionary<string, object>
            {
                ["result"] = resultValue ?? "",
                ["status"] = "success"
            };
        }

        /// <summary>
        /// Trigger cell evaluation
        /// </summary>
        private async Task<object> EvaluateCell(Dictionary<string, object> parameters)
        {
            var sheetName = parameters["sheet"].ToString() ?? "Sheet1";
            var row = Convert.ToInt32(parameters["row"]);
            var column = Convert.ToInt32(parameters["column"]);

            var tcs = new TaskCompletionSource<bool>();
            
            _dispatcherQueue.TryEnqueue(() =>
            {
                try
                {
                    var sheet = _workbook.GetSheet(sheetName) ?? _workbook.Sheets.FirstOrDefault();
                    var cell = sheet?.GetCell(row, column);
                    if (cell != null)
                    {
                        cell.RawValue = cell.RawValue;
                    }
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            
            await tcs.Task;
            return new Dictionary<string, object> { ["status"] = "success" };
        }

        public void Dispose()
        {
            Stop();
            _cancellationTokenSource?.Dispose();
        }
    }

    /// <summary>
    /// IPC Message structure
    /// </summary>
    public class IPCMessage
    {
        public string Command { get; set; } = string.Empty;
        public Dictionary<string, object> Parameters { get; set; } = new();
        public int RequestId { get; set; }
    }
}
