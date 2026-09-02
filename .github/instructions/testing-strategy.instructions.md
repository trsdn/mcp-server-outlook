---
applyTo: "tests/**/*.cs"
---

# Testing Strategy - Quick Reference

## Test Execution

> **Always specify the test project explicitly** to avoid running every test project.

Test projects:

| Project | Covers |
|---|---|
| `tests/OutlookMcp.Core.Tests` | Outlook Core command business logic |
| `tests/OutlookMcp.ComInterop.Tests` | `OutlookDispatcher`, OLE message filter, retained session layer |
| `tests/OutlookMcp.McpServer.Tests` | MCP protocol, generated tools, coverage assertions |
| `tests/OutlookMcp.CLI.Tests` | CLI argument handling, exit codes, daemon |
| `tests/OutlookMcp.Diagnostics.Tests` | Diagnostic helpers |
| `tests/OutlookMcp.SkillGeneration.Tests` | Generated `SKILL.md` and MCP prompt output |

### Everything that runs without Outlook

```powershell
dotnet test --filter "Category!=Integration"
```

### Targeted by Feature trait

```powershell
dotnet test tests\OutlookMcp.ComInterop.Tests\OutlookMcp.ComInterop.Tests.csproj --filter "Feature=OutlookDispatcher"
dotnet test tests\OutlookMcp.McpServer.Tests\OutlookMcp.McpServer.Tests.csproj --filter "Feature=McpProtocol"
dotnet test tests\OutlookMcp.CLI.Tests\OutlookMcp.CLI.Tests.csproj --filter "Feature=CliExitCode"
```

Valid `Feature` values today: `ActionEnums`, `ActionValidation`, `Batch`, `CliExitCode`,
`Configuration`, `DestructiveAnnotations`, `Diag`, `FileLocking`, `McpProtocol`,
`OutlookDispatcher`, `OutlookMcpService`, `OutlookSeed`, `ParameterTransforms`, `PptBatch`,
`PptSession`, `ServiceDaemon`, `ServiceRegistry`, `SessionManager`, `SkillGeneration`,
`StreamJsonRpc`, `VersionCheck`.

`PptBatch` and `PptSession` cover the retained, Outlook-unused session layer in
`src/OutlookMcp.ComInterop/Session/`. They are not Outlook tests.

### Run a specific test by name

```powershell
dotnet test tests\OutlookMcp.Core.Tests\OutlookMcp.Core.Tests.csproj --filter "FullyQualifiedName~TestMethodName"
```

---

## Integration tests need a real Outlook

Anything marked `[Trait("Category", "Integration")]` drives live Outlook COM through
`OutlookInteropRunner`. There is currently **no CI path that runs a single Outlook
operation**: `integration-tests.yml` is wired correctly but reports
`integration-runner-disabled` until the repository variable `ENABLE_OUTLOOK_INTEGRATION_CI` is
set to `true` and a self-hosted Windows runner labelled `outlook` is available. See #31 and
`docs/AZURE_SELFHOSTED_RUNNER_SETUP.md`.

This matters when you write a test: an integration test you add today will compile, will not
run in CI, and will only be exercised by someone running it locally against their own Outlook.
Say so in the PR rather than implying it is verified.

---

## Round-Trip Validation Pattern

**Always verify actual Outlook state after an operation, not just the success flag.**

```csharp
// CREATE -> verify it exists
var createResult = _commands.CreateDraft(subject: "TestSubject", body: "...");
Assert.True(createResult.Success);

var listResult = _commands.List(folder: "drafts");
Assert.Contains(listResult.Items, i => i.Subject == "TestSubject");  // proves it exists

// UPDATE -> verify the change landed
var updateResult = _commands.Update(entryId, subject: "NewSubject");
Assert.True(updateResult.Success);

var readResult = _commands.Read(entryId);
Assert.Equal("NewSubject", readResult.Subject);  // proves the update applied

// DELETE -> verify it is gone
var deleteResult = _commands.Delete(entryId);
Assert.True(deleteResult.Success);

var finalList = _commands.List(folder: "drafts");
Assert.DoesNotContain(finalList.Items, i => i.Subject == "NewSubject");  // proves deletion
```

### Content Replacement Validation (CRITICAL)

For operations that replace content, verify the content was **replaced**, not merged or
appended:

```csharp
// WRONG: only checks the operation completed
var updateResult = _commands.Update(entryId, body: newBody);
Assert.True(updateResult.Success);  // does not prove the body was replaced

// CORRECT: verify the new content is present AND the old content is gone
var updateResult = _commands.Update(entryId, body: newBody);
Assert.True(updateResult.Success);

var readResult = _commands.Read(entryId);
Assert.Equal(newBody, readResult.Body);
Assert.DoesNotContain("OldBody", readResult.Body);

// BETTER: two sequential updates expose merging bugs that a single update hides
_commands.Update(entryId, body: bodyOne);
_commands.Update(entryId, body: bodyTwo);
var readResult = _commands.Read(entryId);
Assert.Equal(bodyTwo, readResult.Body);
Assert.DoesNotContain(bodyOne, readResult.Body);
```

**Why critical:** a real bug in this repository's history had an update method *merging*
content instead of replacing it. Every test passed, because every test only asserted
`Success == true`. The corruption compounded with each update.

**Lesson:** "the operation completed" is not "the operation did the right thing".

---

## Assertion discipline

**Binary assertions only.** Never write an assertion that accepts either outcome:

```csharp
// WRONG: passes whether the feature works or not
Assert.True(result.Success || result.ErrorMessage != null);

// CORRECT
Assert.True(result.Success);
Assert.Null(result.ErrorMessage);
```

If an operation legitimately cannot run without Outlook, skip the test explicitly rather than
softening the assertion. A test that cannot fail is worse than no test, because it reports
green.

---

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Only asserting the success flag | Verify actual Outlook state |
| "Accept both" assertions | Binary assertions only |
| Mocking `Outlook.Application` or `NameSpace` | See ADR-001 - mocked COM proves nothing |
| Mutating the user's real mailbox in a test | Operate on drafts you created, and clean up |
| Missing `Feature` trait | Add one from the list above |
| Leaving an integration test undeclared | Mark `[Trait("Category", "Integration")]` |

---

## When Tests Fail

1. Run individually: `--filter "FullyQualifiedName=Namespace.Class.Method"`
2. Check isolation - does the test depend on state another test left behind?
3. Check assertions - binary, not conditional?
4. Verify Outlook state, not just the success flag
5. If it is a COM failure, check whether Outlook is running at the same elevation as the test
   host. `GetActiveObject` cannot see across integrity levels.

---

## Unit tests

See `docs/ADR-001-NO-UNIT-TESTS.md` for the current, narrowed position. The short version:
tests that mock COM interfaces are prohibited, because the bugs that matter in this repository
(STA affinity, RCW lifetime, Object Model Guard, marshalling) only appear against real
Outlook. Tests for pure logic with no COM dependency - argument parsing, exit-code mapping,
enum-to-string mapping, JSON shaping - are fine and are not what that rule is about.

---

## LLM Integration Tests

**Location**: `llm-tests/`

> **Status (#68):** every scenario in `llm-tests/` targets the PowerPoint surface deleted by
> #26 - charts, tables, ranges, slides, styling. `scripts/Test-LlmRegressionGate.ps1` names six
> of those dead modules as its canonical gate. The harness does not currently test anything
> that exists. #68 tracks the decision to delete it or rewrite it for Outlook. Treat anything
> below as describing the mechanism, not a working suite.

**Purpose**: validate that LLMs correctly use the MCP Server and CLI tools, using
[pytest-aitest](https://github.com/sbroenne/pytest-aitest).

**When to run**: manual and on demand only. It is not part of CI.

```powershell
cd llm-tests
uv sync

uv run pytest -m mcp -v      # MCP Server tests
uv run pytest -m cli -v      # CLI tests
uv run pytest -m aitest -v   # All LLM tests
```

**Prerequisites:**
- `AZURE_OPENAI_ENDPOINT` environment variable
- Windows desktop with classic Outlook for Windows installed and running
- MCP Server built in Release, and `outlookcli` available on PATH

**Configuration overrides:**
- `outlook_mcp_SERVER_COMMAND` overrides the MCP server command
- `OUTLOOK_CLI_COMMAND` overrides the CLI command

See `llm-tests/README.md` for complete documentation.
