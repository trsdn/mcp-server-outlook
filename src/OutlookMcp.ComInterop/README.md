# OutlookMcp.ComInterop

**Low-level COM interop plumbing.** Windows only.

## Overview

This library is the foundation layer for the other `OutlookMcp.*` projects. It owns STA thread
affinity, COM object lifecycle, OLE message filtering, and the named-pipe service protocol.

It contains two clearly separated parts:

| Part | Status | Used by Outlook? |
|---|---|---|
| `OutlookDispatcher`, `ComUtilities`, `OleMessageFilter`, `Progress/`, `ServiceClient/` | Active | Yes |
| The rest of `Session/` (`PptSession`, `PptBatch`, `PptContext`, `SessionManager`, `PptShutdownService`, `ResiliencePipelines`, `IPptBatch`) | **Dormant legacy** | **No** |

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
- **`ResiliencePipelines`** - Polly pipelines for COM retry.
- **`FileAccessValidator`** - path validation, used by attachment save and add.

## Dormant: the retained `Ppt*` session layer

`PptSession`, `PptBatch`, `PptContext`, `SessionManager`, `PptShutdownService`, and `IPptBatch` are
a file-centric session model inherited from the PowerPoint project this repository was forked from.
They open a document, run batched operations against it, and save.

**Nothing calls them.** The PowerPoint command surface they served was deleted in #26, and Outlook
does not use them. They are retained on purpose, as a working, tested reference implementation of a
document-oriented COM session model that a future third COM product could be generalised from. That
generalisation is issue #12.

Two consequences worth stating plainly:

- PowerPoint wording in this part of the codebase is **correct**, not leftover drift. These types
  really are PowerPoint types.
- They should **not** be renamed to `Outlook*`. Doing so would imply Outlook runs on them, which it
  does not.

Note that `PptToolsBase` (in the MCP server project) is a separate matter: it is `Ppt*`-named but the
generated Outlook tools genuinely depend on it. That is real naming debt, also tracked under #12.

## Requirements

- Windows
- .NET 9.0
- For the active Outlook path: the **classic Outlook for Windows desktop app**, installed and
  running. The new Outlook for Windows exposes no COM object model and cannot be automated.

## Platform Support

- Windows x64
- Windows ARM64
- Not Linux or macOS: Office COM is unavailable there.
