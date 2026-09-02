# OutlookMcp.ComInterop

**Low-level COM interop plumbing.** Windows only.

## Overview

This library is the foundation layer for the other `OutlookMcp.*` projects. It owns STA thread
affinity, COM object lifecycle, OLE message filtering, and the named-pipe service protocol.

It contains a single, active Outlook execution path.

| Part | Used by Outlook? |
|---|---|
| `OutlookDispatcher`, `ComUtilities`, `OleMessageFilter`, `Progress/`, `ServiceClient/` | Yes |

## Active: the Outlook execution path

Outlook COM calls run through `OutlookDispatcher`, which owns a single dedicated STA thread and
marshals every call onto it. This is the decision recorded in
[ADR-002](../../docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md).

Outlook is a **shared, user-facing desktop application**, not a file the server opens and closes.
There is no document to create, save, or dispose, so the dispatcher deliberately has no session or
batch concept. It attaches to the running `Outlook.Application` instance and marshals calls to it.

Supporting pieces on this path:

- **`ComUtilities`** - COM object release and safe property access. Every `dynamic` COM object must
  be released in a `finally` block via `ComUtilities.Release(ref obj!)`. This is enforced by
  `scripts/check-com-leaks.ps1` in the pre-commit hook.
- **`OleMessageFilter`** - implements `IMessageFilter` so that busy or rejected **incoming** COM
  calls are retried rather than failing outright. Note the asymmetry: it governs incoming calls
  only. It cannot interrupt an outbound call that is already blocked in the COM runtime, which is
  why dispatcher timeout handling is genuinely hard (issue #19).
- **`FileAccessValidator`** - path validation, used by attachment save and add.

## Requirements

- Windows
- .NET 9.0
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows exposes no COM object model and cannot be automated.

## Platform Support

- Windows x64
- Windows ARM64
- Not Linux or macOS: Office COM is unavailable there.