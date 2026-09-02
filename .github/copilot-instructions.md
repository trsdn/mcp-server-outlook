# GitHub Copilot Instructions - OutlookMcp

> Modular, path-specific instructions for AI coding agents working in this repository.

## Critical files (read these first)

**Always read when working on code:**
- [Critical Rules](instructions/critical-rules.instructions.md) - the mandatory rules (Success flag, COM cleanup, tests, PR process)
- [Architecture Patterns](instructions/architecture-patterns.instructions.md) - dispatcher execution model, command pattern, resource management

**Read based on task type:**
- Adding or fixing commands -> [Outlook COM Interop](instructions/outlook-com-interop.instructions.md)
- Writing tests -> [Testing Strategy](instructions/testing-strategy.instructions.md)
- MCP Server work -> [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- Creating a PR -> [Development Workflow](instructions/development-workflow.instructions.md)
- Fixing bugs -> [Bug Fixing Checklist](instructions/bug-fixing-checklist.instructions.md)

**Less frequently needed:**
- [README Management](instructions/readme-management.instructions.md) - only when updating READMEs
- [Documentation Structure](instructions/documentation-structure.instructions.md) - only when creating docs

---

## What is OutlookMcp?

**OutlookMcp** is a Windows-only toolset for programmatic **Outlook** automation via COM
interop, designed for coding agents and automation scripts. It drives the user's already
running classic Outlook for Windows; it never launches its own instance.

> **OutlookMcp has TWO equal entry points: MCP Server AND CLI.**
> Both are first-class. Every feature, action and parameter must work identically through both.
> When adding or changing features, always verify both are updated. See Rule 24.

**Core layers:**
1. **ComInterop** (`src/OutlookMcp.ComInterop`) - `OutlookDispatcher` (the process-wide STA
   thread that every Outlook call is marshalled onto), the OLE message filter, and COM lifecycle
   helpers.
2. **Core** (`src/OutlookMcp.Core`) - Outlook business logic. `Commands/` contains exactly
   six directories: `Application`, `Attachment`, `Calendar`, `Folder`, `Mail`, and the
   `OutlookInterop` helper that hosts `OutlookInteropRunner`.
3. **Service** (`src/OutlookMcp.Service`) - command routing and registry (in-process for the
   MCP Server, named pipe for the CLI daemon).
4. **CLI** (`src/OutlookMcp.CLI`) - `outlookcli`, for scripting. Equal entry point.
5. **MCP Server** (`src/OutlookMcp.McpServer`) - Model Context Protocol for AI assistants.
   Equal entry point.

**Source generators** (`src/OutlookMcp.Generators*`) - generate CLI commands and MCP tools
from Core interfaces annotated with `[ServiceCategory]` and `[ServiceAction]`. You do not
hand-write tool or command plumbing; you annotate the Core interface and the surface follows.

---

## Quick reference

### Test commands

Integration tests require a real running Outlook and are gated behind a self-hosted runner
(see #31). Locally, run the non-integration suite plus the targeted trait for what you changed.

```powershell
# Everything that runs without Outlook installed
dotnet test --filter "Category!=Integration"

# Targeted by Feature trait
dotnet test --filter "Feature=OutlookDispatcher"
dotnet test --filter "Feature=McpProtocol"
dotnet test --filter "Feature=ServiceDaemon"
```

Valid `Feature` trait values today: `ActionEnums`, `ActionValidation`, `Batch`, `CliExitCode`,
`Configuration`, `DestructiveAnnotations`, `Diag`, `FileLocking`, `McpProtocol`,
`OutlookDispatcher`, `OutlookMcpService`, `OutlookSeed`, `ParameterTransforms`,
`ServiceDaemon`, `ServiceRegistry`, `SkillGeneration`,
`StreamJsonRpc`, `VersionCheck`. Confirm against the source before relying on any of them.

### Code patterns

```csharp
// Core: every Outlook operation goes through OutlookInteropRunner.Execute.
// The runner owns STA marshalling, Application resolution, timeouts and OMG detection.
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
                return new OutlookFolderListResult { Success = true /* ... */ };
            }
            finally
            {
                // Only finally blocks for COM cleanup. No catch returning a failure result.
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

The CLI and MCP surfaces for that method are generated. You do not write them.

---

## Key lessons

**Success flag:** never `Success = true` alongside an `ErrorMessage`. Set `Success = true`
only on the real success path; the `onException` delegate always sets it false.

**No batch or session API.** Outlook has no document to open, save or close - there is one
long-lived running application, which is why every call goes through `OutlookDispatcher`. The
inherited document-session layer was deleted; do not reintroduce a batch API. See
`docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md`.

**Never final-release the shared `Outlook.Application`.** It is the user's single running
instance and its RCW is shared process-wide. Use `ReleaseSharedComObject` for it and
`ReleaseComObject` for everything your call navigated to. See #19.

**Outlook quirks:** collections are 1-based. `Items.Restrict` uses the DASL/Jet dialect, not
SQL. `GetActiveObject` cannot see across integrity levels, so an elevated process cannot
reach an unelevated Outlook. `EntryId` is only meaningful together with `StoreId`.

**Object Model Guard:** Outlook raises a modal prompt for out-of-process callers touching
protected members such as `SenderEmailAddress`, `Recipients` and `MailItem.Send()`. It cannot
be answered programmatically. The runner detects it and reports it distinctly.

**MCP design:** prompts are shortcuts, not tutorials. LLMs already know Outlook concepts and
programming; document what is specific to this server.

**Pre-commit:** search for TODO/FIXME/HACK, delete commented-out code, verify tests, check docs.

**PR review:** check automated review comments (Copilot, GitHub Advanced Security) immediately
and fix them before requesting human review.

**MCP parameter naming:** never use underscores in C# Core interface parameter names. The
generator calls `StringHelper.ToSnakeCase()` on the C# name to produce the MCP snake_case
parameter. Choose a camelCase name that produces the snake_case you want (`entryId` ->
`entry_id`, `storeId` -> `store_id`). If it cannot, use `[FromString("desiredName")]` rather
than an underscore in the C# name.

---

## How path-specific instructions work

GitHub Copilot auto-loads instructions based on the files you are editing:

- `**` (all files) -> [Critical Rules](instructions/critical-rules.instructions.md)
- `src/**/*.cs` -> [Architecture Patterns](instructions/architecture-patterns.instructions.md)
- `src/OutlookMcp.Core/**/*.cs` -> [Outlook COM Interop](instructions/outlook-com-interop.instructions.md)
- `src/OutlookMcp.McpServer/**/*.cs` -> [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- `tests/**/*.cs` -> [Testing Strategy](instructions/testing-strategy.instructions.md)
- `.github/workflows/**/*.yml` -> [Development Workflow](instructions/development-workflow.instructions.md)

---

## Pre-commit hooks

`scripts/pre-commit.ps1` blocks the commit if any check fails:

| # | Check | Script | What it validates |
|---|-------|--------|-------------------|
| 1 | Branch | (inline) | Never commit to `master` directly (Rule 6) |
| 2 | COM leaks | `check-com-leaks.ps1` | COM objects are released in `finally` blocks |
| 3 | Coverage | `CoreCommandsCoverageTests` | Core methods exposed via MCP Server with a matching enum action |
| 4 | Success flag | `check-success-flag.ps1` | Rule 1: never `Success=true` with `ErrorMessage` |
| 5 | CLI settings usage | `check-cli-settings-usage.ps1` | All Settings properties are used in args |
| 6 | CLI workflow test | `Test-CliWorkflow.ps1` | End-to-end CLI smoke test |
| 7 | MCP smoke test | `dotnet test --filter "...SmokeTest..."` | All MCP tools functional |

**Note (#25):** `audit-core-coverage.ps1`, `check-mcp-core-implementations.ps1`,
`check-cli-coverage.ps1` and `check-cli-action-coverage.ps1` were removed. They regex-scraped a
hand-authored `ToolActions.cs` that predates the move to Roslyn source generators (#5/#11) and
no longer exists, so they either false-greened ("0/0 = 100% coverage") or hard-failed on a
missing file. `CoreCommandsCoverageTests` - a reflection-based xUnit test enumerating the live
Outlook Core interfaces and generated enums - replaces their role and is wired into both the
pre-commit hook and CI (`build-cli.yml` / `build-mcp-server.yml`).

**Install hook:**
```powershell
Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit
```

---

## Agent skills (`skills/`)

| Skill | File | Target | Best for |
|-------|------|--------|----------|
| **outlook-cli** | `skills/outlook-cli/SKILL.md` | CLI tool | Coding agents (token-efficient, `--help` discoverable) |
| **outlook-mcp** | `skills/outlook-mcp/SKILL.md` | MCP Server | Conversational AI (rich tool schemas) |

**Build skills from source:**
```powershell
dotnet build -c Release  # Generates SKILL.md, copies references, generates MCP prompts
```

**Guidance architecture (single source of truth):**
- `skills/shared/*.md` is auto-copied to skill references AND auto-generated as MCP prompts
- Skill-based clients (VS Code, Cursor) read `skills/outlook-*/references/`
- MCP-only clients (Claude Desktop) read the auto-generated `[McpServerPrompt]` methods
- **`skills/outlook-*/SKILL.md` are build output. Never edit them.** They are generated from
  `skills/templates/SKILL.cli.sbn` and `skills/templates/SKILL.mcp.sbn`; a Release build silently
  reverts any direct edit. Change the template or `skills/shared/*.md` instead.
- Never create separate prompt files for content that belongs in `skills/shared/`

**Install via npx:**
```powershell
npx skills add trsdn/mcp-server-outlook --skill outlook-cli   # Coding agents
npx skills add trsdn/mcp-server-outlook --skill outlook-mcp   # Conversational AI
```

---

## Architecture patterns

### Command file structure

```
Commands/Mail/
  IMailCommands.cs   # Interface with [ServiceCategory] / [ServiceAction] - drives generation
  MailCommands.cs    # Implementation
```

Split into partial classes by feature domain once a class grows past roughly 15 methods.
One public class per file; file name matches class name.

### Exception handling

`OutlookInteropRunner.Execute` takes an `onException` delegate. That is the single place a
Core command converts an exception into a failed result.

```csharp
// CORRECT: the runner's onException owns failure mapping
return OutlookInteropRunner.Execute(
    "OutlookMailRead",
    (application, session) => { /* ... */ },
    ex => new ActiveMailResult { Success = false, ErrorMessage = ex.Message });

// WRONG: a catch inside the action that returns a failure result.
// It shadows the runner's Object Model Guard classification and loses the real cause.
```

`finally` blocks for COM release are required and are not affected by this rule.

### Service architecture (two equal entry points)

```
MCP Server --> in-process OutlookMcpService --> Core Commands --> OutlookDispatcher --> Outlook COM
CLI --------> CLI daemon (named pipe) ------> Core Commands --> OutlookDispatcher --> Outlook COM
```

Both entry points are first-class. Each hosts its own `OutlookMcpService` instance:
- **MCP Server**: fully in-process, direct method calls, no pipe
- **CLI**: daemon process behind a named pipe (`OutlookMcp-cli-{SID}`) that persists across
  CLI invocations
- **Feature parity**: every action in MCP must exist in the CLI and vice versa
- **Parameter parity**: same parameters, same defaults, same validation

Both funnel into the same single process-wide `OutlookDispatcher` STA thread.
