namespace OutlookMcp.ComInterop;

/// <summary>
/// Constants for Outlook COM interop operations.
/// </summary>
public static class ComInteropConstants
{
    #region Timeouts

    /// <summary>
    /// Default timeout for individual Outlook operations (5 minutes).
    /// Most operations complete in under 30 seconds, but this provides buffer for slow machines
    /// and for large mailboxes where a folder scan can take a while.
    /// </summary>
    public static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Timeout for joining the STA dispatcher thread during shutdown.
    /// Must exceed <see cref="DefaultOperationTimeout"/> so a request that is still running
    /// has a chance to finish before the thread is abandoned.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout = DefaultOperationTimeout + TimeSpan.FromSeconds(15);

    #endregion

    #region Sleep Intervals

    /// <summary>
    /// Delay between file lock acquisition retries (100ms).
    /// </summary>
    public const int FileLockRetryDelayMs = 100;

    #endregion
}
