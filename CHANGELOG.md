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

- **ADR-002: Outlook COM execution model** (#40): recorded the decision to build a purpose-built Outlook STA dispatcher rather than reusing PowerPoint's `PptBatch`/`PptSession` session layer, since Outlook has a single shared always-running `Application` (identified by `entryId`/`storeId`) rather than PowerPoint's per-file open/close model. Documents the fate of `ComInterop/Session/*` (retained for PowerPoint only, deleted alongside the legacy surface in #26) and the naming plan for the new dispatcher.

### Fixed

- **Shared Outlook `Application` COM object could be invalidated process-wide** (#19): `OutlookInteropRunner` called `Marshal.FinalReleaseComObject` on the Outlook `Application` obtained from `GetActiveObject`, which is the user's already-running, shared instance cached per-process by the RCW table. Final-releasing it zeroed the ref-count for every holder in the process, risking `InvalidComObjectException` on subsequent operations. Now released via `Marshal.ReleaseComObject` (a plain ref-count decrement) instead, added a regression test that issues two sequential `Execute()` calls and asserts the second still succeeds.
- **`check-com-leaks.ps1` never scanned Outlook COM files** (#21): the script only flagged leaks in files using PowerPoint's `dynamic` COM pattern, so every Outlook file (which uses strongly-typed `Outlook.*` locals released via `OutlookInteropRunner.ReleaseComObject`) was silently skipped. The script now also detects the Outlook pattern and its release calls.
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
