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

See #19. That is the best-known instance of a more general rule, which is the next section.

---

## Rule: which release call - `ReleaseComObject` or `ReleaseSharedComObject`

`OutlookInteropRunner.ReleaseComObject` is `Marshal.FinalReleaseComObject`, and the CLR's RCW
cache is keyed by the object's `IUnknown` pointer. So "final-release" does not mean "release my
reference" - it means **release everyone's reference to that COM identity, in this process**.

That gives one rule of thumb:

> Final-release is only ever safe for an object **your call navigated to and nobody else holds**.
> Anything Outlook hands back **from a cache** - the same pointer on every access - must use
> `ReleaseSharedComObject`, which is a plain decrement.

| Use `ReleaseComObject` (final) | Use `ReleaseSharedComObject` (decrement) |
|---|---|
| `MAPIFolder` you resolved by path or default-folder role | The shared `Application` (#19) |
| `Items`, `Recipients`, `Attachments`, `Selection` | A parent handed back by a navigation property, e.g. `item.Parent`, `folder.Store` (#122) |
| `MailItem`/`AppointmentItem` fetched by entry id | Everything *under* a `Rules` collection: `Rule`, `RuleConditions`/`RuleActions`, clause slots (#15) |
| The `Rules` collection itself, from `Store.GetRules()`, and the `Store` it came from | Any object you fetch twice in one operation, whatever its type |

`Application.ActiveExplorer()` and `ActiveInspector()` are final-released by existing code and
that has not caused trouble, but they are singletons and sit close to the line. Do not fetch
either twice in one operation.

**The failure mode is not an exception.** A disconnected RCW throws
`InvalidComObjectException`, which is at least catchable. The worse case is that the refcount
reached zero, Outlook freed the object, and the next access goes through a dangling pointer -
an access violation that takes the **whole process** down, with no stack and nothing in the
event log. It is also non-deterministic, so it looks like flakiness.

Four distinct instances have been found in this migration, which is why this is a rule rather
than a footnote:

1. **The shared `Application`** (#19).
2. **The test suite** (#116) - thirteen sites across ten integration test files final-released the
   shared `Application`, so a later test got a wrapper separated from its RCW and the host died
   with `STATUS_STACK_BUFFER_OVERRUN`. It never surfaced as a test failure, only as apparent
   infrastructure flakiness. The convention was documented but not followed.
3. **A parent `MAPIFolder`** obtained from `item.Parent` inside a listing loop (#122). Outlook
   returns the same folder object for every item in it, so final-releasing the parent
   disconnected the very folder the loop was still enumerating, and the test host crashed. Fixed
   by removing the per-item round-trip entirely - the store id is read once from the folder being
   listed - which is the second defence below rather than a different release call.
4. **The whole `Rules` object graph** (#15). `Rules.Item(3)` twice is one object; so are
   `Conditions.Subject` and the `Conditions[n]` that reports the subject slot, and
   `rule.Conditions` fetched before and after a write. Reading a rule's clauses and then
   writing them in the same operation crashed the test host until every `Rule`,
   `RuleConditions`/`RuleActions` and clause slot moved to `ReleaseSharedComObject`.

This applies to **test code exactly as much as to production code** - #116 was entirely in the
test suite, and a crashed test host looks like flakiness rather than like a bug.

**When in doubt, decrement.** An over-retained RCW is collected when the process exits; an
over-released one corrupts an object the user's Outlook is still using.

**A second, cheaper defence, and often the better one:** fetch a child object **once per
operation** and pass it down, rather than re-reading the property. That removes the hazard
instead of handling it, and usually removes a COM round-trip too. #122 took this route - reading
the store id once from the folder being listed rather than per item - and `RuleCommands.Update`
reads a rule's clauses before applying the patch specifically so it never asks Outlook for
`rule.Conditions` a second time. Reach for it before reaching for a different release call.

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

## No session or batch API

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
- [ ] Every `ReleaseComObject` is on an object this call navigated to and nobody else holds -
      anything Outlook hands back from a cache uses `ReleaseSharedComObject`
- [ ] No COM property fetched twice in one operation where the result could be passed down instead
- [ ] No `catch` returning a failure result inside `action`
- [ ] `Success = true` only where `ErrorMessage` is empty
- [ ] `check-com-leaks.ps1` reports 0 leaks
- [ ] `check-success-flag.ps1` passes
- [ ] Any `dynamic` has a comment explaining why
