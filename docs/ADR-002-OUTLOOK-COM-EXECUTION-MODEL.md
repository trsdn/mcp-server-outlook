# ADR-002: Outlook COM Execution Model — Purpose-Built Dispatcher, Not `PptBatch` Reuse

**Status**: Accepted (partially implemented — see [Implementation Status](#implementation-status))
**Date**: 2026-09-01
**Last Updated**: 2026-09-02
**Decision Makers**: Architecture Team
**Stakeholders**: Development Team, Code Reviewers, Contributors
**Related**: #40 (this decision), #12 (epic), #20 (implementation), #19, #29

---

## Context and Problem Statement

The repository currently contains two unrelated COM execution models:

- **Model A — `ComInterop/Session/*`** (inherited from the PowerPoint product): `PptSession`,
  `PptBatch`, `PptContext`, `SessionManager`, `ResiliencePipelines`, `OleMessageFilter`,
  `PptShutdownService`. A long-lived dedicated STA thread pumps a `Channel<Func<Task>>` work
  queue, serially executing operations against one or more open `PowerPoint.Presentation`
  instances. Backed by Polly resilience pipelines, operation timeouts, a named-pipe daemon for
  the CLI, and 8 integration test files.
- **Model B — `OutlookInteropRunner`** (the current Outlook path): every command method is
  `[NoSession]`. Each call to `OutlookInteropRunner.Execute` spawns a **new** STA thread,
  registers `OleMessageFilter`, resolves `Outlook.Application` via `GetActiveObject` (or
  creates one), runs the callback, and tears the thread down. There is no session object, no
  batching across calls, no resilience/retry policy, and — before #19/#20 — a shared-object
  use-after-release bug and no concurrency gate.

We must commit to one execution model before writing #20 (STA dispatcher), #29 (send
confirmation/idempotency), and the rest of the Outlook correctness work in epic #12, because
those issues need a settled place to live.

## Why This Is Not Simply "Port Outlook Onto Model A"

`PptBatch` is shaped around **opening a document and holding a handle to it**:

- Construction takes one or more **file paths**; the constructor's job is to open (or create)
  those specific presentations in a fresh `PowerPoint.Application` it owns.
- The session's lifetime *is* the file's lifetime: `Dispose()` closes the presentations and can
  shut down the `Application` process it created.
- Identity is a **path** (`_presentationPath`, `_allPresentationPaths`), and the batch is
  reasonably one-instance-per-file so independent files can be worked on in parallel via
  separate sessions.

Outlook has none of these properties:

- There is no "document" to open. There is **one shared, already-running `Outlook.Application`**
  that belongs to the interactive user, whether or not this tool is invoked.
- We do not create it, must not create it speculatively, and must never shut it down or
  final-release it (#19) — it is not ours to own.
- The identity primitive for an Outlook item is `EntryID` (+ optional `StoreID`), not a file
  path. There is nothing analogous to "open this file, then operate on it."
- Because the `Application` is a single shared singleton for the whole process (and, via COM,
  effectively for the whole user session), there is exactly **one** session to manage, not one
  per unit of work. `PptBatch`'s per-file multi-instance design solves a problem Outlook doesn't
  have.

Forcing Outlook through `PptBatch`'s shape would mean inventing a fake "path" per operation,
faking open/close semantics around an object that is never actually opened or closed by us, and
carrying PowerPoint-only concepts (`_createNewFile`, `_isMacroEnabled`, per-file process
tracking via `GetWindowThreadProcessId`) that have no Outlook equivalent. That is more
accidental complexity than the moderate duplication of re-solving STA affinity and timeout
handling for a genuinely different resource-lifetime model.

## Decision

**We adopt a purpose-built Outlook dispatcher (an evolution of Model B), not a reuse of
`PptBatch`.**

Specifically:

1. **One process-wide STA dispatcher owns the shared `Outlook.Application`.** Unlike today's
   `OutlookInteropRunner.Execute`, which spins up a new STA thread per call, the dispatcher owns
   a single dedicated STA thread for the lifetime of the process (mirroring `PptBatch`'s
   `Channel<Func<Task>>` work-queue pattern, which *is* worth reusing) and registers
   `OleMessageFilter` exactly once, not per-call. This directly implements #20.
2. **The dispatcher never creates or destroys the shared `Application`.** It resolves it once
   (via `GetActiveObject`, falling back to `Activator.CreateInstance` only if genuinely
   necessary) and holds a live reference for the dispatcher's lifetime. Cleanup uses
   `Marshal.ReleaseComObject` (ref-count decrement), never `FinalReleaseComObject`, consistent
   with the #19 fix already applied in `OutlookInteropRunner.ReleaseSharedComObject`.
3. **Work items are entryId/storeId-addressed, not path-addressed.** Every dispatched operation
   takes whatever identifiers it needs (`entryId`, `storeId`, folder name, etc.) as parameters;
   there is no notion of a session "opening" an item the way `PptBatch` opens a presentation.
4. **Timeouts should cancel or quarantine the in-flight operation rather than abandoning the STA
   thread**, addressing the failure mode in #19 where an abandoned operation could still reach its
   `finally` block and mutate shared COM state after the caller had already moved on. This
   feeds directly into #29's at-most-once requirement for `mail.send`.
   **⚠️ Decided but not yet implemented — see [Implementation Status](#implementation-status).**
5. **Resilience (retry/backoff) is adopted selectively, not wholesale.** `ResiliencePipelines`'
   Polly-based retry policies are reusable as a *library* dependency (they don't assume
   PowerPoint), but the dispatcher does not adopt `PptSession`/`SessionManager`'s multi-session
   lifecycle management, since there is only ever one Outlook session.

## Fate of `ComInterop/Session/*`

- **Retained as-is for PowerPoint.** `PptBatch`, `PptSession`, `PptContext`, `SessionManager`,
  and `PptShutdownService` continue to serve the legacy PowerPoint command surface until that
  surface is removed (epic #11 / issues #23–#26, #34). They are not deleted by this ADR.
- **Not renamed or generalized.** We explicitly reject retrofitting these types into a
  PowerPoint/Outlook-agnostic abstraction — the attempt would either leak PowerPoint-only
  concepts into Outlook call sites or force Outlook's single-shared-instance model through
  multi-instance-per-file abstractions, both of which this ADR found unacceptable.
  `OleMessageFilter` is the one piece of infrastructure both models legitimately share as-is
  (STA message filtering has no PowerPoint- or Outlook-specific shape) and should stay a shared
  `ComInterop` utility.
  When the PowerPoint surface is deleted (#26), `ComInterop/Session/*` should be deleted
  wholesale rather than repurposed — there is no reason to keep `PptBatch`'s file-open/close
  lifecycle around once nothing calls it.

## Naming Plan

- New Outlook dispatcher type: `OutlookMcp.ComInterop.Session.OutlookDispatcher` (or
  `OutlookComDispatcher` if `OutlookDispatcher` collides with a generated name), living beside
  the existing `Session/` folder in `OutlookMcp.ComInterop` since it is genuinely a
  session/dispatch concern, not a `Commands`-layer concern. `OutlookInteropRunner` in
  `OutlookMcp.Core` becomes a thin caller of the dispatcher rather than owning thread lifecycle
  itself.
- No `Ppt*`-prefixed Outlook types are introduced. Existing `Ppt*` types remain scoped to the
  PowerPoint surface and are deleted with it (#26), not renamed for Outlook reuse.
- `OleMessageFilter` keeps its current name and location (`OutlookMcp.ComInterop`) as the one
  shared piece of infrastructure.

## Constraints Satisfied

- ✅ One STA context owns the shared `Outlook.Application`; message filter registered once (#20)
- ✅ Never final-releases the user's `Application` (#19, already fixed in `OutlookInteropRunner`
  ahead of the dispatcher landing, and to be carried into the dispatcher unchanged)
- ⬜ Timeouts cancel/quarantine rather than abandon (feeds #29) — **not yet implemented**; the
  dispatcher currently abandons the *wait*, not the *work*. See below.
- ✅ Multi-step workflows are expressible via `entryId`/`storeId` parameters without requiring a
  file-like session handle to round-trip through the LLM
- ✅ We do not own Outlook's lifetime: no create-on-demand (the untrusted
  `Activator.CreateInstance` fallback was removed in #30 in favor of failing with guidance), no
  shutdown-on-idle

## Implementation Status

| Decision | Status | Where |
|---|---|---|
| 1. Single long-lived STA dispatcher, filter registered once | ✅ Done | `ComInterop/Session/OutlookDispatcher.cs` |
| 2. Never create/destroy the shared `Application` | ✅ Done | `OutlookInteropRunner.ReleaseSharedComObject` |
| 3. `entryId`/`storeId`-addressed work items | ✅ Done | `Commands/Mail/*`, `Commands/Folder/*` |
| 4. Timeout cancels/quarantines in-flight work | ⬜ **Not done** | `OutlookDispatcher.Execute` (#19) |
| 5. Selective resilience adoption | ⬜ Not started | — |

### Why decision 4 is still open

`OutlookDispatcher.Execute` builds a `CancellationTokenSource(timeout)` and applies it to two
things: enqueueing the work item, and `tcs.Task.WaitAsync(...)`. When the timeout fires during
execution, `WaitAsync` throws and the caller receives a `TimeoutException` — but **the queued
delegate keeps running on the STA thread to completion**, including any `finally` blocks that
release COM objects or mutate Outlook state. The `TaskCompletionSource` result is simply
discarded. Cancellation is therefore *observational*, not *effective*.

This is not straightforwardly fixable:

- A blocking cross-apartment COM call cannot be interrupted from another thread. The token is
  never observed because control is inside `Outlook.MailItem.Send()` (or similar), not in
  managed code that could poll it.
- `Thread.Abort` is unsupported on .NET 5+ (`PlatformNotSupportedException`), so the escape
  hatch that made this tractable on .NET Framework is gone.
- `OleMessageFilter` can reject/retry *incoming* calls and is the correct lever for the
  "Outlook is showing a modal dialog" case, but it does not cancel a call we have already made.

Two properties partially mitigate the risk today, and are worth stating explicitly so this is
not mistaken for the pre-#20 bug:

- **No concurrent mutation.** Because all work is serialized onto one STA thread, an abandoned
  operation cannot race a subsequent one — the next work item is head-of-line blocked behind it.
  This is strictly safer than the pre-#20 model, where each call got its own STA thread and
  abandoned threads genuinely could run concurrently against shared COM state.
- **Send is already guarded.** `mail.send` reports `Indeterminate = true` on dispatcher timeout
  rather than `Success = false`, so the one operation with irreversible side effects does not
  silently claim "not sent" while the abandoned delegate completes the send.

The residual problems are that a timeout **stalls the dispatcher** for as long as the abandoned
operation runs (every later caller then times out in the queue-wait path), and that side effects
still land after the caller has given up. A real fix needs a quarantine strategy — e.g. marking
the dispatcher unhealthy after a timeout, and either retiring the STA thread and rebuilding a
fresh one for subsequent work, or refusing further work until the stuck operation returns.
Tracked as the open criterion on #19.

## Consequences

- `OutlookInteropRunner.Execute`'s per-call STA thread spin-up is replaced by dispatching onto
  the single long-lived dispatcher thread — this is the concrete implementation work for #20.
  **Implemented**: `OutlookDispatcher.Shared` owns the thread and the message filter;
  `OutlookInteropRunner.Execute` is now a thin caller.
- `mail.send` confirmation/idempotency (#29) is implemented as dispatcher-level operation
  tracking (an operation token / at-most-once guard keyed off the dispatched work item), not as
  a `PptBatch`-style batch-scoped concept.
  **Implemented**: `MailCommands.Send` requires an explicit `confirm=true` (refused otherwise)
  and accepts an optional caller-supplied `operationId`. Results are cached in-process per
  `operationId` (a `ConcurrentDictionary`, not persisted across restarts) so a caller retrying
  after a timeout with the same `operationId` replays the first attempt's outcome rather than
  re-invoking `MailItem.Send()`. A dispatcher `TimeoutException` during send is surfaced as
  `Indeterminate = true` (distinct from `Success = false`) since the message may have actually
  sent when the timeout fired -- callers must re-check via `mail.read` before deciding whether to
  resend, rather than treating an indeterminate outcome as "definitely not sent". This cache lives
  at the `MailCommands`/Core layer (keyed by caller-provided ID), not inside `OutlookDispatcher`
  itself, since idempotency is a `mail.send`-specific concern, not a generic dispatcher concern
  that every Outlook operation needs.
- The Outlook Object Model Guard (#30) is modelled as a dispatcher-visible outcome (allowed /
  denied / user-cancelled) rather than a swallowed exception, since the dispatcher is now the
  single choke point through which every Outlook COM call passes.
  **Partially implemented**: `OutlookInteropRunner.Execute` now classifies `COMException`s
  consistent with an OMG denial (`E_ABORT`) and surfaces a distinct, actionable
  `InvalidOperationException` instead of a generic COM failure; call sites that read
  security-sensitive properties (`SenderEmailAddress`, `To`/`Cc`/`Bcc`) record which specific
  property reads were blocked in an `AccessDenied` result field rather than returning an
  ambiguous `null`. The untrusted `Activator.CreateInstance` fallback was removed — if no
  running Outlook instance can be reached, the call now fails with guidance instead of creating
  a fresh (more OMG-prone) instance, and an elevation-mismatch (`GetActiveObject` failing because
  this process runs at a different integrity level than Outlook) is detected and reported with a
  specific message. A full dispatcher-level `OperationOutcome` enum (allowed/denied/cancelled)
  threaded through every command's public result type is **not yet done** and remains a
  candidate for a follow-up issue if per-outcome-typed results (rather than
  message-string-based) become necessary.
- Follow-up implementation issues for #12 should be opened against this decision, scoped as:
  (a) introduce `OutlookDispatcher` and migrate `OutlookInteropRunner.Execute` callers onto it
  (**done**, #20),
  (b) add operation cancellation/quarantine on timeout (**open**, remaining criterion on #19),
  (c) wire Guard-outcome modelling (#30) through the dispatcher's return path (**partially done**,
  see above).

## Alternatives Considered

- **Reuse `PptBatch` directly, treating "the Outlook profile" as the analog of "the file
  path".** Rejected: there is exactly one profile/session per process in practice (Outlook does
  not support opening "another mailbox" as a separate `Application`), so `PptBatch`'s
  multi-instance-per-resource design and its own/create/destroy file lifecycle add complexity
  with no corresponding benefit, while its naming and constructors are pervasively
  PowerPoint-shaped (`_isMacroEnabled`, `_createNewFile`, presentation dictionaries).
- **Generalize `ComInterop/Session/*` into an abstract multi-app session layer shared by both
  products.** Rejected for this iteration: the abstraction would need to model "opens a
  document" vs. "attaches to a singleton" as two fundamentally different resource lifecycles,
  which is a larger refactor than the current Outlook correctness work justifies, and the
  PowerPoint surface this layer serves is itself scheduled for deletion (#11/#26). Revisit only
  if a third COM product is added after the PowerPoint surface is gone.
