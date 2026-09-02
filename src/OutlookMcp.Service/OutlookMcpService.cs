using System.IO.Pipes;
using System.Text.Json;
using OutlookMcp.ComInterop.Session;
using OutlookMcp.Core.Commands.Attachment;
using OutlookMcp.Core.Commands.Application;
using OutlookMcp.Core.Commands.Calendar;
using OutlookMcp.Core.Commands.Folder;
using OutlookMcp.Core.Commands.Mail;
using OutlookMcp.Service.Rpc;
using StreamJsonRpc;
using OutlookMcp.Generated;

namespace OutlookMcp.Service;

/// <summary>
/// Main service host for the migration stack.
/// Runs in-process within the host (MCP Server or CLI), accepting commands via named pipe.
/// Exposes the Outlook command surface plus the retained presentation session/batch
/// infrastructure (see ADR-002), which no command category currently routes to.
/// </summary>
public sealed class OutlookMcpService : IDisposable
{
    private readonly SessionManager _sessionManager = new();
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly DateTime _startTime = DateTime.UtcNow;
    private string _pipeName = "";
    private TimeSpan? _idleTimeout;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private bool _disposed;

    // Outlook command instances
    private readonly ApplicationCommands _applicationCommands = new();
    private readonly FolderCommands _folderCommands = new();
    private readonly AttachmentCommands _attachmentCommands = new();
    private readonly MailCommands _mailCommands = new();
    private readonly CalendarCommands _calendarCommands = new();

    public OutlookMcpService()
    {
    }

    public DateTime StartTime => _startTime;
    public int SessionCount => _sessionManager.GetActiveSessions().Count;
    public SessionManager SessionManager => _sessionManager;

    /// <summary>
    /// Runs the service in-process, listening for commands on the named pipe.
    /// This method blocks until shutdown is requested via <see cref="RequestShutdown"/>.
    /// </summary>
    /// <param name="pipeName">The named pipe to listen on.</param>
    /// <param name="idleTimeout">Optional idle timeout. Service shuts down after this duration with no active sessions. Null = no timeout.</param>
    public async Task RunAsync(string pipeName, TimeSpan? idleTimeout = null)
    {
        _pipeName = pipeName;
        _idleTimeout = idleTimeout;
        await RunPipeServerAsync(_shutdownCts.Token);
    }

    public void RequestShutdown() => _shutdownCts.Cancel();

    // Exposed for testing — backoff parameters for pipe server accept loop error recovery
    internal static readonly TimeSpan InitialBackoff = TimeSpan.FromMilliseconds(100);
    internal static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Records client activity to keep the idle timeout monitor alive.
    /// Called by <see cref="Rpc.DaemonRpcTarget"/> on each incoming RPC call.
    /// </summary>
    internal void RecordActivity() => _lastActivityTime = DateTime.UtcNow;

    private async Task RunPipeServerAsync(CancellationToken cancellationToken)
    {
        // Use a semaphore to limit concurrent connections (prevents resource exhaustion)
        using var connectionLimit = new SemaphoreSlim(10, 10);

        // Start idle timeout monitor if configured
        if (_idleTimeout.HasValue)
        {
            _ = Task.Run(() => MonitorIdleTimeoutAsync(cancellationToken), cancellationToken);
        }

        var currentBackoff = InitialBackoff;

        while (!cancellationToken.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = ServiceSecurity.CreateSecureServer(_pipeName);
                await server.WaitForConnectionAsync(cancellationToken);

                // Success — reset backoff
                currentBackoff = InitialBackoff;

                // Record activity on each connection
                _lastActivityTime = DateTime.UtcNow;

                // Capture server for the task
                var clientServer = server;
                server = null; // Prevent disposal in finally - task owns it now

                // Handle client via StreamJsonRpc — replaces hand-rolled JSON protocol
                // with standard JSON-RPC 2.0 over Content-Length-delimited framing.
                _ = Task.Run(async () =>
                {
                    await connectionLimit.WaitAsync(cancellationToken);
                    try
                    {
                        var rpcTarget = new DaemonRpcTarget(this);
                        using var rpc = JsonRpc.Attach(clientServer, rpcTarget);
                        await rpc.Completion; // Waits until client disconnects
                    }
                    finally
                    {
                        connectionLimit.Release();
                        try { if (clientServer.IsConnected) clientServer.Disconnect(); } catch { }
                        await clientServer.DisposeAsync();
                    }
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception)
            {
                // Backoff to prevent CPU spin when errors repeat (e.g. pipe creation failure).
                // Doubles each iteration: 100ms → 200ms → 400ms → … → 5s cap.
                // Resets to 100ms on next successful connection.
                try { await Task.Delay(currentBackoff, cancellationToken); } catch (OperationCanceledException) { break; }
                currentBackoff = TimeSpan.FromMilliseconds(Math.Min(currentBackoff.TotalMilliseconds * 2, MaxBackoff.TotalMilliseconds));
            }
            finally
            {
                if (server != null)
                {
                    try { if (server.IsConnected) server.Disconnect(); } catch (Exception) { /* Cleanup — disconnect may fail if client already disconnected */ }
                    await server.DisposeAsync();
                }
            }
        }
    }

    private async Task MonitorIdleTimeoutAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);

            var hasSessions = _sessionManager.GetActiveSessions().Count > 0;
            if (hasSessions)
            {
                _lastActivityTime = DateTime.UtcNow;
                continue;
            }

            var idleTime = DateTime.UtcNow - _lastActivityTime;
            if (idleTime >= _idleTimeout!.Value)
            {
                RequestShutdown();
                break;
            }
        }
    }

    /// <summary>
    /// Processes a service request directly (in-process, no pipe).
    /// Used by the MCP Server for direct in-process communication.
    /// </summary>
    /// <remarks>
    /// Every remaining command category dispatches synchronously, so this no longer awaits
    /// anything. The <see cref="Task{TResult}"/> signature is kept deliberately: it is the
    /// in-process entry point callers already await, and the asynchrony is expected to return
    /// once dispatcher-backed categories land.
    /// </remarks>
    public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
    {
        try
        {
            // Route command
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            return Task.FromResult(category switch
            {
                "service" => HandleServiceCommand(action),
                "diag" => HandleDiagCommand(action, request),
                "application" => DispatchApplicationSessionless(action, request),
                "attachment" => DispatchAttachmentSessionless(action, request),
                "calendar" => DispatchCalendarSessionless(action, request),
                "folder" => DispatchFolderSessionless(action, request),
                "mail" => DispatchMailSessionless(action, request),
                "session" => HandleSessionCommand(action, request),
                _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" }
            });
        }
        catch (Exception ex)
        {
            // Include type name so callers can distinguish exception kinds (GitHub #482, Bug 5)
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" });
        }
    }

    // === SERVICE COMMANDS ===

    private ServiceResponse HandleServiceCommand(string action)
    {
        return action switch
        {
            "ping" => new ServiceResponse { Success = true },
            "shutdown" => HandleShutdown(),
            "status" => HandleStatus(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown service action: {action}" }
        };
    }

    private ServiceResponse HandleShutdown()
    {
        _shutdownCts.Cancel();
        return new ServiceResponse { Success = true };
    }

    private ServiceResponse HandleStatus()
    {
        var status = new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = _sessionManager.GetActiveSessions().Count,
            StartTime = _startTime
        };
        return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(status, ServiceProtocol.JsonOptions) };
    }

    // === LEGACY POWERPOINT SESSION COMMANDS ===

    // === DIAG COMMANDS ===

    private static ServiceResponse HandleDiagCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "ping" => new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    action = "ping",
                    message = "pong",
                    timestamp = DateTime.UtcNow.ToString("o")
                }, ServiceProtocol.JsonOptions)
            },
            "echo" => HandleDiagEcho(request),
            "validate-params" => HandleDiagValidateParams(request),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown diag action: {action}" }
        };
    }

    private static ServiceResponse HandleDiagEcho(ServiceRequest request)
    {
        Dictionary<string, JsonElement>? args = null;
        if (!string.IsNullOrEmpty(request.Args))
            args = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(request.Args, ServiceProtocol.JsonOptions);

        if (args == null || !args.TryGetValue("message", out var messageEl) || messageEl.ValueKind == JsonValueKind.Null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Parameter 'message' is required for echo" };
        }

        var message = messageEl.GetString()!;
        string? tag = null;
        if (args.TryGetValue("tag", out var tagEl) && tagEl.ValueKind != JsonValueKind.Null)
            tag = tagEl.GetString();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                action = "echo",
                message,
                tag
            }, ServiceProtocol.JsonOptions)
        };
    }

    private static ServiceResponse HandleDiagValidateParams(ServiceRequest request)
    {
        Dictionary<string, JsonElement>? args = null;
        if (!string.IsNullOrEmpty(request.Args))
            args = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(request.Args, ServiceProtocol.JsonOptions);

        if (args == null || !args.TryGetValue("name", out var nameEl) || nameEl.ValueKind == JsonValueKind.Null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Parameter 'name' is required for validate-params" };
        }

        var count = args.TryGetValue("count", out var countEl) && countEl.ValueKind == JsonValueKind.Number ? countEl.GetInt32() : 0;
        string? label = args.TryGetValue("label", out var labelEl) && labelEl.ValueKind != JsonValueKind.Null ? labelEl.GetString() : null;
        var verbose = args.TryGetValue("verbose", out var verboseEl) && verboseEl.ValueKind != JsonValueKind.Null && verboseEl.GetBoolean();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                action = "validate-params",
                parameters = new
                {
                    name = nameEl.GetString(),
                    count,
                    label,
                    verbose
                }
            }, ServiceProtocol.JsonOptions)
        };
    }

    // === SESSION COMMANDS ===

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown session action: {action}" }
        };
    }

    private ServiceResponse HandleSessionCreate(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        var fullPath = Path.GetFullPath(args.FilePath);

        if (File.Exists(fullPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"File already exists: {fullPath}. Use session open to open an existing legacy presentation."
            };
        }

        var extension = Path.GetExtension(fullPath);
        if (!string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".pptm", StringComparison.OrdinalIgnoreCase))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Invalid file extension '{extension}'. legacy session create supports .pptx and .pptm only."
            };
        }

        try
        {
            // Legacy PowerPoint path: use the combined create+open which starts PowerPoint only once
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSessionForNewFile(fullPath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);

            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = fullPath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    private ServiceResponse HandleSessionOpen(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        try
        {
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSession(args.FilePath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = args.FilePath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    private ServiceResponse HandleSessionClose(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var args = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);
        var closed = _sessionManager.CloseSession(request.SessionId, save: args?.Save ?? false);

        return closed
            ? new ServiceResponse { Success = true }
            : new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
    }

    private ServiceResponse HandleSessionSave(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var batch = _sessionManager.GetSession(request.SessionId);
        if (batch == null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
        }

        // Legacy PowerPoint path: check if the presentation process is still alive before attempting save
        if (!batch.IsPowerPointProcessAlive())
        {
            _sessionManager.CloseSession(request.SessionId, save: false, force: true);
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Legacy PowerPoint process for session '{request.SessionId}' has died. Session has been closed. Please create a new session."
            };
        }

        batch.Save();
        return new ServiceResponse { Success = true };
    }

    private ServiceResponse HandleSessionList()
    {
        var sessions = _sessionManager.GetActiveSessions()
            .Select(s => new
            {
                sessionId = s.SessionId,
                filePath = s.FilePath,
                isPowerPointVisible = _sessionManager.IsPowerPointVisible(s.SessionId),
                activeOperations = _sessionManager.GetActiveOperationCount(s.SessionId),
                canClose = _sessionManager.GetActiveOperationCount(s.SessionId) == 0
            })
            .ToList();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessions, count = sessions.Count }, ServiceProtocol.JsonOptions)
        };
    }



    // === GENERATED DISPATCH ===

    // All command routing uses ServiceRegistry.*.DispatchToCore() generated methods.

    // See ServiceRegistry.*.Dispatch.g.cs for the generated code.







    private static ServiceResponse WrapResult(string? dispatchResult)

    {

        return dispatchResult == null

            ? new ServiceResponse { Success = true }

            : new ServiceResponse { Success = true, Result = dispatchResult };

    }



    private ServiceResponse DispatchApplicationSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Application.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.Application.DispatchToCore(_applicationCommands, action, request.Args));
    }

    private ServiceResponse DispatchAttachmentSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Attachment.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.Attachment.DispatchToCore(_attachmentCommands, action, request.Args));
    }

    private ServiceResponse DispatchCalendarSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Calendar.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.Calendar.DispatchToCore(_calendarCommands, action, request.Args));
    }

    private ServiceResponse DispatchFolderSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Folder.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.Folder.DispatchToCore(_folderCommands, action, request.Args));
    }

    private ServiceResponse DispatchMailSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Mail.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.Mail.DispatchToCore(_mailCommands, action, request.Args));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _shutdownCts.Cancel();
        _sessionManager.Dispose();
        _shutdownCts.Dispose();
    }
}

// === ARGUMENT TYPES (Session only - all other args are now generated in ServiceRegistry) ===

// Session
public sealed class SessionOpenArgs
{
    public string? FilePath { get; set; }
    public bool Show { get; set; }
    public int? TimeoutSeconds { get; set; }
}
public sealed class SessionCloseArgs { public bool Save { get; set; } }
