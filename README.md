# mcp-server-outlook

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9-blue.svg)](https://dotnet.microsoft.com/download/dotnet/9.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](#requirements)

**Outlook automation MCP server and CLI for classic Outlook on Windows.**

This repository was bootstrapped from a mature Office-automation MCP architecture and has since
been rebuilt around Outlook. The inherited presentation surface - 33 command domains and the
document-session layer that backed them - has been removed; what remains is Outlook-only.

- the tool surface is 5 tools and 30 operations, identical through the MCP server and the CLI
- there is no session or batch concept: every COM call is marshalled onto a single STA thread
- the long-term goal is full Outlook COM coverage with MCP, CLI, and VS Code extension parity

## Current reality

What is true right now:

- the copied architecture is mature and reusable: `Core`, `Service`, `ComInterop`, generators, CLI, MCP server, VS Code extension, skills, and tests
- the semantic surface is Outlook-only: 6 tools, 48 operations, identical through the MCP server and the CLI (see [FEATURES.md](FEATURES.md))
- there is no session or batch concept: every COM call is marshalled onto a single STA thread owned by `OutlookDispatcher` (see [ADR-002](docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md))
- **no Outlook behaviour is verified by CI, and none ever will be.** There is no self-hosted Windows runner with classic Outlook and there is no plan to provide one (#31, closed as not planned), so every correctness claim about Outlook rests on a local run against a real profile
- this repo should be treated as an Outlook rebuild on top of reusable plumbing, not as a large search-and-replace exercise

## Implemented Outlook surface

The repository exposes **6 tools with 48 operations**, generated into both the MCP server and the
CLI from the same `[ServiceCategory]` interfaces:

| Tool | Operations |
|---|---|
| `mail` | 22 |
| `folder` | 10 |
| `calendar` | 6 |
| `contact` | 5 |
| `attachment` | 4 |
| `application` | 1 |

See [FEATURES.md](FEATURES.md) for the full action list.

## Planned next Outlook slice

The seed has since been extended to attachments, calendar, folder mutation and contacts. The
remaining gaps are:

- richer server-side mail search, replacing the current client-side scan (#42)
- a paging cursor for large result sets (#43)
- tasks (`TaskItem`)
- `application` window state: the active explorer and inspector

The single most important gap is not a feature: it is verification, and it is a permanent one.
Nothing enforces that Outlook behaviour was checked before a merge, because no CI job can check it.
Run the integration tests locally against a real profile, and say plainly when you have not.

## Architecture direction

The migration keeps the proven generator-driven shape of the original project where it still makes sense:

- `src\OutlookMcp.Core` defines command interfaces and implementations
- `src\OutlookMcp.Service` handles routing and IPC-backed orchestration
- `src\OutlookMcp.Generators.Mcp` and `src\OutlookMcp.Generators.Cli` generate MCP and CLI surfaces
- `src\OutlookMcp.McpServer` hosts the MCP server
- `src\OutlookMcp.CLI` provides the command-line surface
- `vscode-extension` packages the user-facing VS Code integration

The biggest architectural difference from the inherited baseline is the execution model:

- the baseline was file-centric and document-centric, with a per-document session
- Outlook is application-, folder-, and item-centric
- Outlook behaves like a single shared desktop application instance

That is why the document-session layer was deleted rather than renamed. See [ADR-002](docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md).

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

The inherited presentation command families are **gone** as of #26: `slide`, `shape`, `text`, `chart`,
`animation`, `transition`, `slideshow`, `master`, `slideimport`, `customshow`, `background`,
`pagesetup`, `shapealign`, `placeholder`, `headerfooter`, `design`, `image`, `media`, `export`,
`smartart`, `vba`, and the rest of the 33 inherited domains. Neither the MCP server nor the CLI
exposes them.

## Important naming note

This repo is `mcp-server-outlook` and the solution and projects use `OutlookMcp.*` throughout.
Historical references to the inherited product remain only in `CHANGELOG.md` and the ADRs,
where they explain why decisions were made.

## Immediate migration priorities

1. ~~Define the Outlook-first command taxonomy.~~ Done.
2. ~~Rebuild COM and service abstractions around Outlook objects.~~ Done (#20, ADR-002).
3. ~~Generate Outlook CLI and MCP surfaces from that taxonomy.~~ Done (#23, #26).
4. ~~Stand up CI that can actually run Outlook (#31).~~ Not planned. Outlook is verified locally against a real profile, or not at all.
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
