using System.Text.Json;
using OutlookMcp.Service;

namespace OutlookMcp.McpServer.ServiceBridge;

/// <summary>
/// Bridge that holds the in-process OutlookMcp Service for direct method calls.
/// No named pipe — MCP tools call the service directly (same process).
/// </summary>
public static class ServiceBridge
{
    private static readonly SemaphoreSlim _initLock = new(1, 1);
    private static Service.OutlookMcpService? _service;

    /// <summary>
    /// JSON serializer options for deserializing service responses.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = ServiceProtocol.JsonOptions;

    /// <summary>
    /// Ensures the in-process OutlookMcp Service is created.
    /// Called automatically on first request.
    /// </summary>
    public static async Task<bool> EnsureServiceAsync(CancellationToken cancellationToken = default)
    {
        if (_service != null)
        {
            return true;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            if (_service != null)
            {
                return true;
            }

            _service = new Service.OutlookMcpService();
            return true;
        }
        catch (Exception)
        {
            return false;
        }
        finally
        {
            _initLock.Release();
        }
    }

    /// <summary>
    /// Sends a command to the OutlookMcp Service directly (in-process, no pipe).
    /// </summary>
    public static async Task<ServiceResponse> SendAsync(
        string command,
        object? args = null,
        int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        if (!await EnsureServiceAsync(cancellationToken))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = "Failed to start OutlookMcp Service in-process."
            };
        }

        var request = new ServiceRequest
        {
            Command = command,
            Args = args != null ? JsonSerializer.Serialize(args, JsonOptions) : null
        };

        // Apply timeout if specified
        if (timeoutSeconds.HasValue)
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds.Value));
            try
            {
                return await _service!.ProcessAsync(request);
            }
            catch (OperationCanceledException) when (cts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
            {
                return new ServiceResponse
                {
                    Success = false,
                    ErrorMessage = $"Operation timed out after {timeoutSeconds} seconds."
                };
            }
        }

        return await _service!.ProcessAsync(request);
    }

    /// <summary>
    /// Disposes the in-process OutlookMcp Service.
    /// Must be called when the MCP server process exits.
    /// </summary>
    public static void Dispose()
    {
        var service = Interlocked.Exchange(ref _service, null);
        service?.Dispose();
    }
}