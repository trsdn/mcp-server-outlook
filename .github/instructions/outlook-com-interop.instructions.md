---
applyTo: "src/OutlookMcp.Core/**/*.cs"
---

# Outlook COM Interop Patterns

> **The single execution model for every Outlook operation in this repository.**
> Authoritative background: `docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md`.

---

## The one pattern

Every Outlook command in `src/OutlookMcp.Core/Commands/` goes through
`OutlookInteropRunner.Execute`. There are no exceptions and no alternative path.

```csharp
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
public OutlookFolderListResult ListDefault()
{
    return OutlookInteropRunner.Execute(
        "OutlookFolderListDefault",
        (application, session) =>
        {
            Outlook.MAPIFolder? folder = null;
            try
            {
                folder = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                return new OutlookFolderListResult
                {
                    Success = true,
                    // ... populate from folder ...
                };
            }
            finally
            {
                OutlookInteropRunner.ReleaseComObject(ref folder);
            }
        },
        ex => new OutlookFolderListResult
        {
            Success = false,
            ErrorMessage = $"Failed to read Outlook default folders: {ex.Message}"
        });
}
```

Three arguments, all required:

| Argument | Purpose |
|---|---|
| `operationName` | A stable, unique string. Shows up in dispatcher diagnostics and timeout messages. Name it `Outlook<Domain><Verb>`, e.g. `OutlookMailRead`, `OutlookFolderListChildren`. |
| `action` | `Func<Outlook.Application, Outlook.NameSpace, TResult>`. Runs on the STA thread. Both COM objects are handed to you already resolved. |
| `onException` | `Func<Exception, TResult>`. The **only** place a Core command turns an exception into a failed result. |

---

## What `Execute` already does for you

Do not re-implement any of this inside `action`:

- **Marshals onto the STA thread.** `OutlookDispatcher.Shared` owns one process-wide STA
  thread. Outlook's object model is apartment-threaded; calling it from a pool thread is
  undefined behaviour.
- **Resolves the running Outlook.** Uses `GetActiveObject` on the `Outlook.Application`
  ProgID. It deliberately does **not** fall back to `Activator.CreateInstance` (see #30) - a
  freshly created instance is not the user's trusted session and is more likely to trip the
  Object Model Guard.
- **Produces actionable failures** for "Outlook not installed", "Outlook not running", and
  "this process is elevated and Outlook is not" (COM cannot see across integrity levels).
- **Detects Object Model Guard denials** and surfaces them as a distinct, explanatory error
  rather than an opaque `COMException`.
- **Applies `ComInteropConstants.DefaultOperationTimeout`.**
- **Releases the per-call `NameSpace`, and never final-releases the shared `Application`.**

---

## Rule: never final-release the shared `Application`

`Outlook.Application` is the user's single already-running instance. The CLR caches one RCW
for it per process. `Marshal.FinalReleaseComObject` zeroes that refcount for **every holder
in the process**, not just your call, so a later operation gets a disconnected RCW.

```csharp
// WRONG: breaks every other operation in the process
Marshal.FinalReleaseComObject(application);

// CORRECT: the runner handles the shared Application itself.
// For a shared object you obtained yourself, use the decrementing variant:
OutlookInteropRunner.ReleaseSharedComObject(ref application);
```

`ReleaseComObject` (final-release) is correct for objects **your call created or navigated
to**: folders, items, `Items` collections, `Recipients`, `Attachments`, `Selection`,
`Explorer`, `Inspector`. Those are per-call and nobody else holds them. See #19.

---

## Rule: strongly-typed interop, not `dynamic`

Core references `Microsoft.Office.Interop.Outlook` and aliases it:

```csharp
using Outlook = Microsoft.Office.Interop.Outlook;
```

Use `Outlook.MailItem`, `Outlook.MAPIFolder`, `Outlook.NameSpace`, `Outlook.OlDefaultFolders`
and friends. Strong typing gives you compile-time checking and correct release semantics via
the generic `ReleaseComObject<T>(ref T?)` overload.

`dynamic` is permitted only where the object's runtime type genuinely varies - most commonly
an `object` pulled from `Selection` or `Items` that may be a `MailItem`, `MeetingItem`,
`ReportItem` or something else entirely. Prefer a type test over `dynamic`:

```csharp
object? item = selection[1];
try
{
    if (item is Outlook.MailItem mail) { /* ... */ }
}
finally
{
    OutlookInteropRunner.ReleaseComObject(ref item);
}
```

If you must use `dynamic`, document why on the line above it - `check-dynamic-casts.ps1`
inspects these.

---

## Rule: try-finally for release, never try-catch

Every COM object your `action` touches must be released in a `finally` block.

```csharp
// CORRECT
Outlook.MAPIFolder? folder = null;
Outlook.Items? items = null;
try
{
    folder = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
    items = folder.Items;
    // ...
}
finally
{
    OutlookInteropRunner.ReleaseComObject(ref items);
    OutlookInteropRunner.ReleaseComObject(ref folder);
}
```

Release in reverse acquisition order: children before parents.

Do **not** add a `catch` inside `action` that returns a failure result. That is what
`onException` is for, and duplicating it there loses the runner's Object Model Guard
classification. See Rule 1b and Rule 22 in `critical-rules.instructions.md`.

The narrow exceptions that remain legitimate:

- `catch { continue; }` when iterating a collection where individual items may be
  inaccessible (a corrupt item should not fail the whole listing).
- Catching around a single optional property read whose absence is meaningful and handled.
- `catch (COMException ex) when (<specific HRESULT>)` for a genuinely different code path.

---

## Rule: `Success` must match reality

`Success = true` implies `ErrorMessage` is null or empty. Always. Set `Success = true` only
on the actual success path inside `action`; `onException` always sets `Success = false`.
`check-success-flag.ps1` enforces this in the pre-commit hook. See Rule 1.

---

## The Object Model Guard

Outlook raises a modal security prompt for out-of-process callers touching protected members:
`SenderEmailAddress`, `Recipients`, `AddressEntry.Address`, `MailItem.Send()`, and others. If
nobody answers the dialog, Outlook aborts the call.

You cannot answer it programmatically, and this server must not try. The runner already
detects the resulting `COMException` shapes and returns an explanatory error. When you add an
operation that touches a protected member, say so in its XML documentation so the tool
description warns the caller.

---

## Retained but dormant: `ComInterop/Session/*`

`PptSession`, `PptBatch`, `PptContext` and `IPptBatch` still exist under
`src/OutlookMcp.ComInterop/Session/`. **No Outlook code path calls them.** They are a
product-neutral document-session layer retained deliberately; ADR-002 records the decision and
`src/OutlookMcp.ComInterop/README.md` explains why they keep their current names.

Do not introduce a batch API into an Outlook command. Outlook has no document to open, save
or close - there is one long-lived running application, which is exactly why the dispatcher
model replaced the batch model.

---

## Outlook API gotchas

- **Collections are 1-based.** `items[1]` is the first item, and `items[0]` throws.
- **`Items.Restrict` and `Items.Find` use the DASL/Jet query dialect**, not SQL and not LINQ.
  Date literals must be formatted with `Format(...)` semantics Outlook understands.
- **`Items.Sort` must be applied before enumeration** to have any effect, and re-sorting
  invalidates any index you were holding.
- **`GetActiveObject` fails across integrity levels.** An elevated process cannot see an
  unelevated Outlook. The runner reports this specifically; do not paper over it.
- **`MailItem.Send()` is asynchronous from the caller's point of view.** It hands off to
  Outbox; delivery is not confirmed by the call returning.
- **`EntryId` is not stable across stores.** An item moved between stores gets a new one, so
  always carry `StoreId` alongside it.

---

## Before writing new COM code

Search other open-source projects for a working example first. Do not search this repository
for one. [NetOffice](https://github.com/NetOfficeFw/NetOffice) has strongly-typed wrappers for
the entire Outlook object model and is the best available reference for correct call shapes
and release semantics. Real VBA and C# samples prevent the usual mistakes: 1-based indices,
missed releases, and protected members that trip the Object Model Guard.

Validate against the
[Outlook VBA reference](https://learn.microsoft.com/office/vba/api/overview/outlook) before
adding any dependency. Use the Outlook COM API for anything it supports.

---

## Checklist

Before committing a change under `src/OutlookMcp.Core/`:

- [ ] Goes through `OutlookInteropRunner.Execute` with a unique `operationName`
- [ ] Every COM object released in a `finally`, children before parents
- [ ] Shared `Application` never final-released
- [ ] No `catch` returning a failure result inside `action`
- [ ] `Success = true` only where `ErrorMessage` is empty
- [ ] `check-com-leaks.ps1` reports 0 leaks
- [ ] `check-success-flag.ps1` passes
- [ ] Any `dynamic` has a comment explaining why
