# Changelog

All notable changes to OutlookMcp (PowerPoint MCP Server) will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Changed

- **BREAKING — Project and package rename `PptMcp.*` → `OutlookMcp.*`** (#5): every project under `src/` and `tests/`, the solution file, all assembly names, root namespaces, `using` directives, NuGet package IDs, build scripts, and CI workflows were renamed so this repository publishes only under its own identity.
  - NuGet package IDs: `PptMcp.McpServer` → `OutlookMcp.McpServer`, `PptMcp.CLI` → `OutlookMcp.CLI`, `PptMcp.Core` → `OutlookMcp.Core`, `PptMcp.ComInterop` → `OutlookMcp.ComInterop`
  - dotnet tool commands: `pptcli` → `outlookcli`, `mcp-ppt` → `mcp-outlook`
  - npm packages: `ppt-mcp-skill` → `outlook-mcp-skill`, `ppt-cli-skill` → `outlook-cli-skill`
  - Release artifacts: `PptMcp-<version>.vsix` → `OutlookMcp-<version>.vsix`, `ppt-skills-v<version>.zip` → `outlook-skills-v<version>.zip`
  - Skill generation now writes to `skills/outlook-mcp` and `skills/outlook-cli`; the stale `skills/ppt-mcp` and `skills/ppt-cli` directories were removed
  - Repository self-references updated from `trsdn/mcp-server-ppt` to `trsdn/mcp-server-outlook`

### Added

- **Ported `folder.resolve-path`, `folder.list-items`, `mail.set-subject`, `mail.set-body`, `mail.set-recipients` from the orphaned `feature/outlook-parity-slices` branch** (#28): that branch was cut before the `PptMcp` → `OutlookMcp` rename (#5) landed on `master` and was never merged, leaving genuinely useful capabilities stranded. Rather than merging or rebasing the whole branch — which also carries untested Contacts CRUD and Application explorer/inspector context work — only the two smallest, self-contained, highest-value slices were cherry-picked and re-ported onto current `OutlookMcp.*` namespaces: `folder.resolve-path`/`folder.list-items` (new mailbox folder navigation, not previously exposed) and `mail.set-subject`/`mail.set-body`/`mail.set-recipients` (draft editing, a prerequisite called out by #28 for future draft-editing (#15) and meeting-creation (#14) work since `Recipients` mutation was otherwise unavailable). The Contacts and Application explorer/inspector slices were deliberately **not** ported in this pass — they need new result types and additional test coverage and are tracked separately. Added `[SkippableFact]` smoke tests for all five new actions in `OutlookSeedSmokeTests`, following the existing integration-test-only pattern (Rule 30).

- **Detect and report classic-vs-new Outlook for Windows** (#35): the new Outlook for Windows (the modern packaged Mail & Calendar replacement) has no COM object model and cannot be automated by this server, but previously a missing `Outlook.Application` ProgID always produced a generic "Outlook not installed" message, even when new Outlook was present and running. Added `OutlookInstallationDetector` (a pure, zero-COM-dependency utility using `Type.GetTypeFromProgID` and new-Outlook package registry checks) that distinguishes `NotInstalled`, `ClassicDesktop`, `NewOutlookOnly`, and `Unknown`, plus a same-process-integrity-level check via `WindowsPrincipal`. `application.get-status` now reports `outlookFlavor`/`processElevated`, and the not-installed error path in `OutlookInteropRunner` surfaces a flavor-specific, actionable message (e.g. telling the user new Outlook is installed but unsupported, vs. classic Outlook installed but not running). Added `outlookcli diag outlook` to surface the same detection standalone. Updated `README.md`'s Requirements section to explicitly call out that classic Outlook desktop (not new Outlook for Windows) is required.

- **ADR-002: Outlook COM execution model** (#40): recorded the decision to build a purpose-built Outlook STA dispatcher rather than reusing PowerPoint's `PptBatch`/`PptSession` session layer, since Outlook has a single shared always-running `Application` (identified by `entryId`/`storeId`) rather than PowerPoint's per-file open/close model. Documents the fate of `ComInterop/Session/*` (retained for PowerPoint only, deleted alongside the legacy surface in #26) and the naming plan for the new dispatcher.

- **Serialize all Outlook COM access behind a single long-lived STA dispatcher** (#20, P0): previously every `OutlookInteropRunner.Execute` call spawned its own short-lived STA thread, registering/revoking `OleMessageFilter` on every single call — this could allow overlapping COM re-entrancy between operations and paid the thread-spin-up cost repeatedly. Added `OutlookDispatcher` (`OutlookMcp.ComInterop.Session`), a process-wide singleton exposing a generic `Execute<T>(operationName, operation, timeout)` that queues work onto one dedicated STA thread via a **bounded** `Channel<Func<Task>>` (capacity 32, blocking on full — giving explicit back-pressure per the issue's acceptance criteria, unlike PowerPoint's unbounded `PptBatch` channel). `OleMessageFilter.Register()`/`Revoke()` now run exactly once for the dispatcher's whole lifetime instead of per call. `OutlookInteropRunner.Execute` now delegates to `OutlookDispatcher.Shared` instead of owning per-call thread lifecycle; its public signature and all 7 existing call sites (`ApplicationCommands`, `MailCommands`, `CalendarCommands`, `AttachmentCommands`, `FolderCommands`) are unchanged. Added unit tests asserting overlapping concurrent callers are serialized onto a single thread, that a thrown exception propagates without wedging the dispatcher, and that a timed-out operation surfaces `TimeoutException` while leaving the dispatcher usable afterward. Note: as with `PptBatch`, a timeout does not interrupt an in-flight COM call already running on the STA thread — true cancellation/quarantine of stuck operations is tracked separately (#29).

- **`mail.send` confirmation gate and at-most-once retry safety** (#29, P0): sending mail is a destructive, one-way action, but `mail.send` previously fired `MailItem.Send()` unconditionally on any call, and a client-side timeout left the caller unable to tell whether the message actually sent — retrying blindly risked a duplicate send. `MailCommands.Send` now requires an explicit `confirm=true` and is refused with a clear error otherwise. Added an optional `operationId` parameter: results are cached in-process per `operationId` (a `ConcurrentDictionary`, not persisted across restarts), so a caller retrying with the same `operationId` after a timeout replays the first attempt's outcome instead of re-invoking `MailItem.Send()`. Added `MailSendResult.Indeterminate`: if the dispatcher's `TimeoutException` fires while a send may already be in flight, the result now reports `Indeterminate = true` (distinct from `Success = false`) with guidance to re-check via `mail.read` before deciding whether to resend, rather than looking like a definite failure. Added unit tests covering the confirmation gate.

### Fixed

- **Outlook Object Model Guard denials were unmodelled and silently swallowed** (#30, P0, security): the Outlook Object Model Guard (OMG) raises a modal security prompt for out-of-process, untrusted callers touching protected members (`SenderEmailAddress`, `Recipients`, `MailItem.Send()`, etc.); a denied/unanswered prompt previously looked identical to any other COM failure, and `MailCommands.SafeGet`'s bare `catch { return null; }` made a blocked `SenderEmailAddress` read indistinguishable from "this mail has no sender" — precisely the pattern Rule 22 forbids. Also, `OutlookInteropRunner.GetOrCreateApplication` fell back to `Activator.CreateInstance` when no running Outlook instance could be found via `GetActiveObject`, deliberately creating an untrusted instance that is *more* likely to trigger OMG and conflicts with Outlook's single-instance model. Removed the `Activator.CreateInstance` fallback entirely — a missing running instance now fails with actionable guidance, distinguishing a plain "Outlook not running" from an elevation mismatch (this process running elevated while Outlook runs unelevated, which makes `GetActiveObject` fail with `MK_E_UNAVAILABLE`-shaped errors) via `OutlookInstallationDetector.IsCurrentProcessElevated()`. Added `OutlookInteropRunner.IsObjectModelGuardDenial` to classify `COMException`s consistent with an OMG denial (`E_ABORT`) and surface a distinct, actionable error instead of a generic COM failure. `MailCommands`' security-sensitive property reads (`SenderEmailAddress`, `To`, `Cc`, `Bcc`) now record which specific properties were blocked in a new `AccessDenied` result field on `ActiveMailResult`/`MailSummaryInfo`, rather than returning an ambiguous `null`. Extended `outlookcli diag outlook` with an `objectModelGuard` section documenting the OMG posture and elevation-mismatch risk. Documented the (partial) OMG posture in `docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md`. A full dispatcher-level allowed/denied/cancelled outcome enum threaded through every command is not yet done and remains a candidate follow-up.
- **16/16 Outlook integration tests silently reported `Passed` with zero assertions executed** (#22): `OutlookSeedSmokeTests` guarded every fact with `if (!EnsureOutlookAvailable()) { return; }`, so on any machine or CI runner without a running classic Outlook desktop instance, the suite reported 16/16 green having exercised nothing — violating the repo's "tests must fail loudly, never silent" rule and masking #16's finding that no Outlook CI runner exists. Switched every guarded fact to `[SkippableFact]` (via the `Xunit.SkippableFact` package) and changed `EnsureOutlookAvailable()` to throw `SkipException` instead of returning `false`, so the same runs now correctly report `Skipped`. Added `scripts/check-outlook-tests-not-all-skipped.ps1`, wired into the CI Outlook job, which fails the build if every Outlook test in a TRX result skipped — a 100% skip rate on a runner that is supposed to have Outlook now fails the job instead of manufacturing false confidence.
- **`calendar` and `attachment` tools under-declared `destructiveHint`, `mail` over-declared it** (#18): action-dispatch tools (one MCP tool, many actions via an `action` enum) cannot be described by a single tool-level boolean. `calendar` declared `Destructive = false` while exposing `delete-appointment`/`update-appointment`; `attachment` declared `Destructive = false` while exposing `add`/`remove`; `mail` declared `Destructive = true` while exposing read-only `list`/`read`/`search`. Added a per-action `[ServiceAction(Destructive = ...)]` override, and the generator now computes each tool's `[McpServerTool(Destructive = ...)]` hint as true if ANY exposed action is destructive. `calendar`'s description no longer claims to manage appointments "safely". Added a regression test asserting every affected action's resolved classification and the generated tool-level hints.
- **Shared Outlook `Application` COM object could be invalidated process-wide** (#19): `OutlookInteropRunner` called `Marshal.FinalReleaseComObject` on the Outlook `Application` obtained from `GetActiveObject`, which is the user's already-running, shared instance cached per-process by the RCW table. Final-releasing it zeroed the ref-count for every holder in the process, risking `InvalidComObjectException` on subsequent operations. Now released via `Marshal.ReleaseComObject` (a plain ref-count decrement) instead, added a regression test that issues two sequential `Execute()` calls and asserts the second still succeeds.
- **`check-com-leaks.ps1` never scanned Outlook COM files** (#21): the script only flagged leaks in files using PowerPoint's `dynamic` COM pattern, so every Outlook file (which uses strongly-typed `Outlook.*` locals released via `OutlookInteropRunner.ReleaseComObject`) was silently skipped. The script now also detects the Outlook pattern and its release calls.
- **MCP `ServerInstructions` taught LLMs the legacy PowerPoint file/session workflow** (#24): the server-wide instructions sent to every connecting client told the model to unlock legacy session-gated tools via `file(action:'open')` → `session_id` → `file(action:'close')`, described the product as a "migration surface", and omitted `calendar` entirely. Rewritten to describe only the Outlook surface (`application`, `folder`, `mail`, `attachment`, `calendar`), the `entryId`/`storeId` identity model, when to use active-item targeting vs. explicit `entryId`, and destructive-action safety expectations.
- **Release workflow published under foreign package identities** (#5): removed the temporary `if: false` guard on the `publish` job now that every registry target is Outlook-owned.
- **MCP registry `server.json` version was never updated**: the version-rewrite regex matched a non-existent `Trsdn.PptMcp.McpServer` identifier and silently did nothing, so a stale version would have been published.
- **Scriban 6.6.0 broke every restore**: the templating engine behind SKILL.md generation carried a critical advisory (GHSA-5wr9-m6jw-xx44, patched in 7.0.0) and `NuGetAudit` treats it as an error, so `dotnet restore` failed outright. Bumped to 7.2.6; generated SKILL.md output is byte-identical.
- **Dependency review rejected the GitHub Copilot CLI license**: added an explicit `allow-dependencies-licenses` entry, since GitHub's proprietary terms cannot be expressed as an SPDX identifier.
- **Vulnerable transitive dependency in the agent lockfile**: `@github/copilot` resolved to 1.0.4, which is affected by GHSA-9ccr-r5hg-74gf (arbitrary command execution via `core.fsmonitor`). Refreshed to 1.0.80 within the existing `@github/copilot-sdk@0.1.32` range.
- **CI never ran on any branch**: all workflows filtered on `main`, but this repository's default branch is `master`, so build, CodeQL, dependency-review, and integration-test workflows were silently inert. `build-mcp-server.yml` was also missing a `pull_request` trigger entirely.
- **NuGet propagation check never succeeded**: the readme poll used a mixed-case package ID, which the lowercase-only flat-container API always answers with 404, wasting the full 30-minute retry window on every release.

### Added

- Official source-side Copilot SDK agent client under `src\OutlookMcp.Agent`, including local planner tests and documentation for the agent architecture
- Dedicated documentation for the evaluation framework and the archetype/reference pipeline
- **33 PowerPoint MCP tools with 204 operations** for comprehensive PowerPoint automation via COM interop
- **Slide management** (7 ops) — list, read, create, duplicate, move, delete, apply-layout
- **Shape operations** (17 ops) — add, move, resize, fill, line, shadow, rotation, z-order, grouping, copy between slides, connectors, merge shapes (union/combine/fragment/intersect/subtract)
- **Text editing** (6 ops) — get/set text, find, replace, format (font, size, bold, italic, color, alignment)
- **Charts** (5 ops) — create, get info, set title, set type, delete
- **Slide Tables** (9 ops) — create, read, write cells, add/delete rows and columns, merge cells
- **Animations** (4 ops) — list, add, remove, clear effects
- **Transitions** (3 ops) — get, set, remove slide transitions
- **Design/Themes** (4 ops) — list designs, apply themes, get theme colors, list color schemes
- **Images** (1 op) — insert with position and size control
- **Speaker Notes** (3 ops) — get, set, clear
- **Sections** (4 ops) — list, add, rename, delete presentation sections
- **Hyperlinks** (4 ops) — add, read, remove, list
- **Slideshow** (4 ops) — start, stop, navigate, get status
- **Slide Masters** (1 op) — list masters and layouts
- **Export** (4 ops) — PDF, slide images (PNG), video (MP4), print
- **VBA Macros** (5 ops) — list, view, import, delete, run
- **Media** (3 ops) — insert audio/video, get media info
- **Window Management** (4 ops) — get info, minimize, restore, maximize
- **File Validation** (1 op) — test file accessibility
- **Document Properties** (2 ops) — get/set title, author, subject, etc.
- **Comments** (4 ops) — list, add, delete, clear slide comments
- **Placeholders** (2 ops) — list placeholders, set placeholder text
- **Slide Background** (3 ops) — get info, set solid color, reset to master
- **Headers & Footers** (2 ops) — get/set footer text, slide numbers, date
- **SmartArt** (2 ops) — get diagram info, add nodes
- **Shape Alignment** (2 ops) — align and distribute shapes on slides
- **Custom Shows** (3 ops) — list, create, delete custom slide shows
- **Page Setup** (2 ops) — get/set slide size and orientation
- **Slide Import** (1 op) — import slides from another .pptx file
- **Tags** (3 ops) — custom metadata on slides and shapes
- **MCP Server** — Model Context Protocol server for AI assistants (GitHub Copilot, Claude, ChatGPT)
- **CLI** (`outlookcli`) — Command-line interface for scripting and coding agents
- **COM interop** — Uses PowerPoint's native COM API for 100% safe automation
- **Session management** — Shared sessions between MCP Server and CLI
- **Parameter validation** — All required string parameters validated before COM execution
- **COM resource safety** — All COM objects released in finally blocks to prevent leaks
