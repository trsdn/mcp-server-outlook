# OutlookMcp Tests

OutlookMcp is now an Outlook COM automation server. The current product surface is 5 Outlook tools with 30 operations: mail, calendar, folder, attachment, and application.

No Outlook behavior is verified by hosted CI today. Real Outlook integration tests require a self-hosted Windows runner with classic Outlook installed, running, and signed in. That runner does not exist yet.

## Quick Start

```powershell
# Build and run CI-safe tests
dotnet build --nologo -v q
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"

# Run all tests locally only when the machine has the required desktop Office setup
dotnet test

# Run Outlook smoke tests manually on a Windows desktop with classic Outlook running
dotnet test tests\OutlookMcp.Core.Tests --filter "Feature=OutlookSeed"

# Run retained dormant session-layer tests only when modifying that legacy infrastructure
dotnet test tests\OutlookMcp.ComInterop.Tests --filter "Feature=PptBatch|Feature=PptSession|Feature=SessionManager"
```

## Documentation

For broader test philosophy and repository rules, see:

- [No Unit Tests ADR](../docs/ADR-001-NO-UNIT-TESTS.md)
- [Outlook COM Execution Model ADR](../docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md)
- [Testing Strategy](../.github/instructions/testing-strategy.instructions.md)
- [Critical Rules](../.github/instructions/critical-rules.instructions.md)

Some instruction files still contain historical PowerPoint guidance while the Outlook migration continues. Follow the current code and this README for the active Outlook surface.

## Test Architecture

```text
tests\
|-- OutlookMcp.Core.Tests\          # Core result contracts and Outlook smoke tests
|-- OutlookMcp.McpServer.Tests\     # MCP protocol, tool metadata, and coverage checks
|-- OutlookMcp.CLI.Tests\           # CLI daemon, batch, diagnostics, and validation tests
|-- OutlookMcp.ComInterop.Tests\    # Shared COM infrastructure and dormant legacy session layer
|-- OutlookMcp.Diagnostics.Tests\   # Manual diagnostics, currently empty or historical
`-- OutlookMcp.SkillGeneration.Tests\ # Skill markdown quality checks

llm-tests\                          # Manual LLM behavior tests
```

## Test Categories

| Category | Purpose | Requirements | CI status |
| -------- | ------- | ------------ | --------- |
| Unit | Pure .NET behavior, serialization, generated metadata, validation | .NET SDK | CI-safe |
| Integration | Protocol or process integration, plus Outlook smoke tests | Depends on test | Only CI-safe subsets should be assumed |
| OutlookSeed | Manual Outlook behavior smoke coverage | Classic Outlook running on Windows | Not verified by hosted CI |
| PptBatch, PptSession, SessionManager | Dormant retained presentation-session infrastructure | Desktop Office setup | Only relevant when modifying that layer |
| LLM tests | Manual AI behavior validation | Azure OpenAI config and local tools | Manual only |

## Outlook Smoke Tests

`OutlookMcp.Core.Tests\Integration\Outlook\OutlookSeedSmokeTests.cs` exercises real Outlook commands when classic Outlook is available.

These tests:

- Require Windows and classic Outlook.
- Require Outlook to be running and signed in.
- Use real mailbox state.
- Can create, mutate, or delete real Outlook items as part of smoke coverage.
- Are not a substitute for a dedicated self-hosted CI runner.

Run them manually:

```powershell
dotnet test tests\OutlookMcp.Core.Tests --filter "Feature=OutlookSeed"
```

## Coverage and Metadata Tests

`CoreCommandsCoverageTests` is the key CI-safe guard that checks every `[ServiceCategory]` interface has generated actions and action-string mappings.

```powershell
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

This verifies generated surface parity, not Outlook runtime behavior.

## Feature-Specific Tests

Useful filters in the current tree:

```powershell
dotnet test --filter "Feature=ActionEnums"
dotnet test --filter "Feature=ActionValidation"
dotnet test --filter "Feature=Configuration"
dotnet test --filter "Feature=McpProtocol"
dotnet test --filter "Feature=OutlookDispatcher"
dotnet test --filter "Feature=OutlookMcpService"
dotnet test --filter "Feature=OutlookSeed"
dotnet test --filter "Feature=ServiceDaemon"
dotnet test --filter "Feature=SkillGeneration"
```

When changing the dormant retained session layer, use its specific filters:

```powershell
dotnet test tests\OutlookMcp.ComInterop.Tests --filter "Feature=PptBatch"
dotnet test tests\OutlookMcp.ComInterop.Tests --filter "Feature=PptSession"
dotnet test tests\OutlookMcp.ComInterop.Tests --filter "Feature=SessionManager"
```

## When to Run Which Tests

| Scenario | Command |
| -------- | ------- |
| Documentation-only change | `dotnet build --nologo -v q` |
| Generated action or Core interface change | `dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"` |
| CLI routing or daemon change | `dotnet test tests\OutlookMcp.CLI.Tests` |
| Outlook command behavior change | Manual `Feature=OutlookSeed` on Windows with classic Outlook running, plus targeted tests you add |
| Dormant session-layer change | Target `PptBatch`, `PptSession`, or `SessionManager` tests |
| Skill guidance change | `dotnet test tests\OutlookMcp.SkillGeneration.Tests` |
| LLM behavior validation | `.\scripts\Test-LlmRegressionGate.ps1` if that script is current for the Outlook surface |

## LLM Tests

The `llm-tests\` project is manual. Before relying on it, inspect the prompts and expected behavior to make sure they reference the current Outlook command surface.

```powershell
Set-Location llm-tests
uv sync
uv run pytest -m aitest -v
```

Prerequisites:

- `AZURE_OPENAI_ENDPOINT` environment variable.
- Windows desktop when tests call local MCP or CLI tools.
- Any required local model or service credentials.

## Key Principles

- Test the active Outlook command surface, not deleted presentation features.
- Do not claim hosted CI verifies Outlook behavior until a self-hosted Windows Outlook runner exists.
- Treat tests that touch Outlook as real mailbox automation.
- Prefer targeted tests for the changed layer.
- Keep coverage checks derived from `[ServiceCategory]` interfaces so new tools cannot be forgotten.

## Getting Help

- Test failures: read the failing test output first.
- Outlook issues: confirm classic Outlook is installed, running, signed in, and at the same elevation level as the test process.
- Generated action issues: inspect Core service interfaces and generated metadata.
- Session-layer issues: remember that `Ppt*` infrastructure is retained but dormant for Outlook commands.
