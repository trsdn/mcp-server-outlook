# mcp-server-outlook

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9-blue.svg)](https://dotnet.microsoft.com/download/dotnet/9.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](#requirements)

**Outlook-first MCP server migration built on the architecture of `mcp-server-outlook`.**

This repository is no longer intended to remain a PowerPoint product. It is the active Outlook migration target bootstrapped from the local `mcp-server-outlook` baseline.

Today the repository is still in transition:

- the repo identity is now Outlook-first
- the internal codebase still contains many inherited `OutlookMcp` names
- many packages, skills, binaries, tests, and docs still reflect PowerPoint-era naming and behavior
- the long-term goal is full Outlook COM coverage with MCP, CLI, and VS Code extension parity

## Current reality

What is true right now:

- the copied architecture is mature and reusable: `Core`, `Service`, `ComInterop`, generators, CLI, MCP server, VS Code extension, skills, eval, and tests
- the semantic surface is still overwhelmingly PowerPoint-shaped
- Outlook behavior has to be rebuilt around mailbox, folder, item, draft, and meeting workflows rather than presentation files
- this repo should be treated as an Outlook rebuild on top of reusable plumbing, not as a large search-and-replace exercise

## Implemented Outlook seed

The repository now contains a first real Outlook seed integrated into Core, Service, MCP generation, and build validation:

- `application.get-status`
- `folder.list-default`
- `mail.read-active`
- `mail.list`
- `mail.search`
- `mail.create-draft`
- `mail.reply`
- `mail.reply-all`
- `mail.forward`
- `mail.send`
- `attachment.list`
- `attachment.save`

This is intentionally small, but it proves the first Outlook-specific path through the inherited architecture.

## Planned next Outlook slice

The next slice should extend that seed toward:

- deeper current folder selection and folder traversal
- mail read and inspect beyond the active item
- recipient, subject, and body editing for existing drafts
- attachment list, add, and save

After that, the next major domains are:

- calendar
- contacts
- follow-up / task workflows if needed

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

### Remove or replace outright

These PowerPoint-centric families should not survive as Outlook concepts:

- `slide`
- `shape`
- `animation`
- `transition`
- `slideshow`
- `master`
- `slideimport`
- `customshow`
- `background`
- `pagesetup`
- `shapealign`
- `placeholder`
- `headerfooter`
- large parts of `design`, `image`, `media`, and `export` as currently modeled

## Important naming note

This repo is now `mcp-server-outlook`, but several inherited names are still present during migration:

- solution and projects still use `OutlookMcp.*`
- package ids and skill names still use inherited `ppt-*` forms
- some docs and examples still refer to PowerPoint
- some public metadata still points at the original PowerPoint lineage until the Outlook surfaces are in place

That cleanup is intentional work still to be completed, not hidden compatibility debt.

## Immediate migration priorities

1. Define the Outlook-first command taxonomy.
2. Rebuild COM and service abstractions around Outlook objects.
3. Generate Outlook CLI and MCP surfaces from that taxonomy.
4. Rewire skills and extension UX to those surfaces.
5. Replace PowerPoint-specific tests, docs, examples, and evals.
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
