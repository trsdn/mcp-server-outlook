namespace OutlookMcp.Service.Rpc;

/// <summary>
/// Server-side RPC target that delegates incoming JSON-RPC calls to <see cref="OutlookMcpService.ProcessAsync"/>.
/// One instance is attached per pipe connection via <c>JsonRpc.Attach(stream, target)</c>.
/// </summary>
internal sealed class DaemonRpcTarget : IOutlookDaemonRpc
{
    private readonly OutlookMcpService _service;

    public DaemonRpcTarget(OutlookMcpService service)
    {
        _service = service;
    }

    /// <inheritdoc />
    public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request)
    {
        _service.RecordActivity();
        return await _service.ProcessAsync(request);
    }
}
