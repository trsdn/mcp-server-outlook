# Timeout Implementation Guide

## Overview

OutlookMcp currently uses `OutlookDispatcher` for active Outlook commands. It serializes all Outlook COM work onto one dedicated STA thread and applies a timeout to both queueing and execution.

**Key Features:**

- Active Outlook operations use `OutlookDispatcher.Shared.Execute`.
- The dispatcher has a bounded queue to provide back-pressure for overlapping callers.
- The timeout covers waiting for a queue slot and running on the STA thread.
- The default active Outlook timeout comes from `ComInteropConstants.DefaultOperationTimeout`.
- Timeout failures are reported as operation failures by the command result path.

---

## Core Implementation

### Constants

`ComInteropConstants.DefaultOperationTimeout` is the shared default timeout used by Outlook commands.

```csharp
public static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromMinutes(5);
```

### Active Outlook Dispatcher

`OutlookDispatcher` is the active execution path.

```csharp
public T Execute<T>(string operationName, Func<T> operation, TimeSpan timeout)
```

Behavior:

1. Write the operation to a bounded channel.
2. If the queue is full until timeout, throw a `TimeoutException` that names dispatcher queue pressure.
3. Run the operation on the single STA thread.
4. If execution does not complete by timeout, throw a `TimeoutException` that names the running operation.

### OutlookInteropRunner

Active Outlook Core commands call through `OutlookInteropRunner.Execute`, which dispatches to `OutlookDispatcher`:

```csharp
return OutlookDispatcher.Shared.Execute(operationName, () =>
{
    // Resolve Outlook.Application and Outlook.NameSpace.
    // Run the command operation.
}, ComInteropConstants.DefaultOperationTimeout);
```

This means active Outlook commands do not take a per-command timeout parameter today.

---

## Enhanced Result Types

Result types inherit common success and guidance fields where applicable.

```csharp
public abstract class ResultBase
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? FilePath { get; set; }
    public List<string>? SuggestedNextActions { get; set; }
    public Dictionary<string, object>? OperationContext { get; set; }
    public bool IsRetryable { get; set; } = true;
    public string? RetryGuidance { get; set; }
}
```

For Outlook commands, prefer actionable guidance that mentions the real likely cause: classic Outlook not running, elevation mismatch, Object Model Guard prompt, dispatcher queue pressure, or item identity changes.

---

## Usage Patterns

### Pattern 1: Active Outlook Commands

Use `OutlookInteropRunner.Execute` for Outlook COM work. Do not create ad hoc STA threads per command.

```csharp
return OutlookInteropRunner.Execute(
    "application.get-status",
    (application, session) =>
    {
        return new OutlookApplicationStatusResult { Success = true };
    },
    ex => new OutlookApplicationStatusResult
    {
        Success = false,
        ErrorMessage = ex.Message
    });
```

### Pattern 2: Timeout Error Guidance

When surfacing dispatcher timeouts, include guidance that fits Outlook:

- Check whether classic Outlook is responsive.
- Check whether an Outlook security or modal dialog is waiting for user input.
- Avoid launching many overlapping Outlook operations.
- Retry only when the operation is safe to repeat.
- For `mail send`, use `operationId` to avoid duplicate sends on retry.

---

## Operation-Specific Timeout Recommendations

| Operation Type | Path | Recommendation |
| -------------- | ---- | -------------- |
| Mail, calendar, folder, attachment, application | Active Outlook dispatcher | Use the shared default unless adding a designed timeout option. |
| Many overlapping CLI or MCP requests | Active Outlook dispatcher | Expect queue back-pressure. Avoid parallel Outlook mutation. |
| Send mail | Active Outlook dispatcher | Use `confirm=true`; use `operationId` for retry safety. |

---

## Stderr Logging

For active Outlook timeout work, prefer messages and result context that name the Outlook operation, for example `mail.send` or `folder.list-default`.

---

## Integration Checklist

### For New Outlook Core Commands

- [ ] Route COM work through `OutlookInteropRunner.Execute`.
- [ ] Keep work short and deterministic where possible.
- [ ] Return result objects with accurate `Success` and `ErrorMessage` values.
- [ ] Add `SuggestedNextActions` for likely Outlook failures.
- [ ] Consider retry safety, especially for destructive operations.

### For MCP Tools

- [ ] Return JSON results for business errors.
- [ ] Preserve generated action metadata.
- [ ] Do not expose deleted presentation tools.

### For CLI Commands

- [ ] Keep CLI and MCP action names and parameters in sync.
- [ ] Surface timeout errors as JSON in quiet/scripted paths.
- [ ] Include actionable guidance without claiming the operation is isolated from Outlook.

---

## Testing Strategy

### Active Outlook Runtime Behavior

Manual only today:

```powershell
dotnet test tests\OutlookMcp.Core.Tests --filter "Feature=OutlookSeed"
```

This requires classic Outlook running on Windows. Hosted CI does not verify Outlook behavior yet.

### Active Metadata and Surface Coverage

CI-safe:

```powershell
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

---

## Troubleshooting

### Symptom: All Outlook operations time out

Likely causes:

1. Classic Outlook is hung or waiting on a modal dialog.
2. An Outlook Object Model Guard prompt is waiting for user input.
3. Too many overlapping operations are queued.
4. Outlook and the calling process are running at different elevation levels.

Solutions:

1. Bring Outlook to the foreground and look for prompts.
2. Restart classic Outlook.
3. Retry with fewer concurrent operations.
4. Run Outlook and the server or CLI at the same elevation level.

### Symptom: Queue timeout before operation starts

Likely cause: too many overlapping Outlook operations in flight.

Solution: serialize callers or reduce parallel requests. Outlook is a single shared desktop app and should not be treated as an isolated multi-worker backend.

### Symptom: Retry could duplicate an action

Likely cause: the client timed out after Outlook may already have completed the operation.

Solution: do not blindly retry destructive actions. For `mail send`, pass an `operationId` so the command can answer repeated attempts from cached send state when possible.

---

## Future Enhancements

Potential improvements:

1. Per-operation timeout configuration for active Outlook commands.
2. Better timeout metrics for dispatcher queue pressure versus execution stalls.
3. More specific result guidance for Object Model Guard prompts.
4. Self-hosted Windows Outlook integration runner for CI.

---

## Related Documentation

- [Outlook COM Execution Model](ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md)
- [Development Workflow](DEVELOPMENT.md)
- [Testing README](../tests/README.md)

---

**Last Updated:** 2026-09-02