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

### Fixed

- **Release workflow published under foreign package identities** (#5): removed the temporary `if: false` guard on the `publish` job now that every registry target is Outlook-owned.
- **MCP registry `server.json` version was never updated**: the version-rewrite regex matched a non-existent `Trsdn.PptMcp.McpServer` identifier and silently did nothing, so a stale version would have been published.
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
