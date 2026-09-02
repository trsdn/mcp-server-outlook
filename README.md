# mcp-server-outlook

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9-blue.svg)](https://dotnet.microsoft.com/download/dotnet/9.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](#requirements)

**Outlook-first MCP server migration built on the architecture of `mcp-server-outlook`.**

This repository is no longer intended to remain a PowerPoint product. It is the active Outlook migration target bootstrapped from the local `mcp-server-outlook` baseline.

Today the repository is still in transition:

- the repo identity is now Outlook-first
- the tool surface is now Outlook-only: the 33 inherited PowerPoint command domains were deleted in #26
- the internal codebase still contains inherited `OutlookMcp` and `Ppt*` type names
- some docs, packaging and tooling still reflect PowerPoint-era naming
- the long-term goal is full Outlook COM coverage with MCP, CLI, and VS Code extension parity

## Current reality

What is true right now:

- the copied architecture is mature and reusable: `Core`, `Service`, `ComInterop`, generators, CLI, MCP server, VS Code extension, skills, and tests
- the semantic surface is Outlook-only: 5 tools, 30 operations, identical through the MCP server and the CLI (see [FEATURES.md](FEATURES.md))
- the retained `ComInterop/Session/*` layer is dormant COM plumbing kept deliberately; nothing calls it, and Outlook does not use it (see [ADR-002](docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md))
- **no Outlook behaviour is verified by CI.** Integration tests need a self-hosted Windows runner with classic Outlook installed, which does not exist yet, so correctness claims rest on local runs only
- this repo should be treated as an Outlook rebuild on top of reusable plumbing, not as a large search-and-replace exercise

## Implemented Outlook surface

The repository exposes **5 tools with 30 operations**, generated into both the MCP server and the
CLI from the same `[ServiceCategory]` interfaces:

| Tool | Operations |
|---|---|
| `mail` | 16 |
| `calendar` | 5 |
| `folder` | 4 |
| `attachment` | 4 |
| `application` | 1 |

See [FEATURES.md](FEATURES.md) for the full action list.

## Planned next Outlook slice

The seed has since been extended to attachments and calendar. The remaining gaps are:

- richer server-side mail search, replacing the current client-side scan (#42)
- a paging cursor for large result sets (#43)
- contacts
- follow-up / task workflows if needed

The single most important gap is not a feature: it is verification. See #31.

## Architecture direction

The migration keeps the proven generator-driven shape of the original project where it still makes sense:

- `src\OutlookMcp.Core` defines command interfaces and implementations
- `src\OutlookMcp.Service` handles routing and IPC-backed orchestration
- `src\OutlookMcp.Generators.Mcp` and `src\OutlookMcp.Generators.Cli` generate MCP and CLI surfaces
- `src\OutlookMcp.McpServer` hosts the MCP server
- `src\OutlookMcp.CLI` provides the command-line surface
- `vscode-extension` packages the user-facing VS Code integration

The biggest architectural change is the session model:

- PowerPoint is file-centric and presentation-centric
- Outlook is application-, folder-, and item-centric
- Outlook also behaves much more like a single shared desktop application instance

That means current `file`, `slide`, `shape`, and presentation-session assumptions cannot simply be renamed.

## Keep / adapt / remove guidance

### Keep or strongly reuse

- named-pipe IPC and host/service split
- generator pipeline
- response/result DTO patterns
- build and packaging scaffolding
- integration-test structure
- extension packaging pipeline

### Adapt heavily

- `ComInterop` session and batch abstractions
- `Service` dispatch categories
- CLI help and command taxonomy
- MCP tool families and descriptions
- shared skills and marketplace UX
- docs, examples, eval scenarios, and smoke tests

### Removed

The PowerPoint-centric command families are **gone** as of #26: `slide`, `shape`, `text`, `chart`,
`animation`, `transition`, `slideshow`, `master`, `slideimport`, `customshow`, `background`,
`pagesetup`, `shapealign`, `placeholder`, `headerfooter`, `design`, `image`, `media`, `export`,
`smartart`, `vba`, and the rest of the 33 inherited domains. Neither the MCP server nor the CLI
exposes them.

## Important naming note

This repo is now `mcp-server-outlook`, but several inherited names are still present during migration:

- solution and projects still use `OutlookMcp.*`
- `Ppt*`-prefixed infrastructure remains, including `PptToolsBase`, which the generated Outlook tools depend on
- the retained `ComInterop/Session/*` types are still `Ppt*`-named; they should **not** be renamed to `Outlook*`, because Outlook does not use them (see #12)
- some docs and examples still refer to PowerPoint

That cleanup is intentional work still to be completed, not hidden compatibility debt.

## Immediate migration priorities

1. ~~Define the Outlook-first command taxonomy.~~ Done.
2. ~~Rebuild COM and service abstractions around Outlook objects.~~ Done (#20, ADR-002).
3. ~~Generate Outlook CLI and MCP surfaces from that taxonomy.~~ Done (#23, #26).
4. **Stand up CI that can actually run Outlook (#31).** Nothing else is trustworthy until this exists.
5. Rewire skills and extension UX to those surfaces.
6. Perform the final coordinated rename of inherited `OutlookMcp` internals.

## Requirements

- Windows
- **Classic Outlook for Windows desktop app installed and running** (the app registering the `Outlook.Application` COM ProgID). The new Outlook for Windows (the modern, packaged replacement for Mail & Calendar) has **no COM object model** and cannot be automated by this server; if only new Outlook is present, `application.get-status` / `outlookcli diag outlook` will report `NewOutlookOnly` and every Outlook action will fail with an actionable message telling you to install or switch to classic Outlook.
- Desktop automation context available

## Validation status

The repository now has a compiling Outlook seed surface and generator-backed MCP exposure for that seed.

What is validated right now:

- solution build passes with the Outlook seed in place
- MCP discovery and service wiring can now include Outlook-first tool families

What is not yet validated:

- real Outlook smoke tests
- end-to-end CLI and extension workflows against the new Outlook families
