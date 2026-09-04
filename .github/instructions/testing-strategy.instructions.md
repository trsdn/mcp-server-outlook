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
| `tests/OutlookMcp.ComInterop.Tests` | `OutlookDispatcher`, OLE message filter |
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
`OutlookDispatcher`, `OutlookMcpService`, `OutlookSeed`, `ParameterTransforms`,
`ServiceDaemon`, `ServiceRegistry`, `SkillGeneration`,
`StreamJsonRpc`, `VersionCheck`.

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

## A red test over `send` actually sends

Rule 29 requires you to watch the test fail before you implement. On a *refusal* — "this send must
be blocked" — the red run performs the very thing the feature exists to prevent, because the block
is not there yet. That is inherent to TDD on a sending surface, not carelessness, and it will
happen again on `mail.send`, `mail.reply-all`, `mail.forward` and
`calendar.create-appointment` with `sendInvitation`.

**Address it so it cannot reach a person.** Use a recipient in a reserved TLD — `.invalid`
(RFC 2606) — or a domain you control and monitor. Never a real colleague, never a distribution
list, and never an address that merely looks unroutable.

```csharp
// The red run of this test WILL send. .invalid can never be delivered.
recipientTo: "blocked@definitely-not-allowed.invalid"
```

**Then clean up.** The message reaches Sent Items even when delivery fails, and a `finally` that
only deletes the *draft* will not catch it — the draft is gone, and what is left is a sent item with
a new entry ID. Check Sent Items and Outbox after any red run of a send test.

**Say so in the test's own doc comment**, so the next person running it against a broken build is
warned before they run it rather than after.

This is worth the ceremony because the alternative is worse: softening the test so it does not send
would make it a test that cannot fail, which is the one thing
[Assertion discipline](#assertion-discipline) forbids.

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
| A red run of a "must be refused" send test actually sending | Reserved `.invalid` recipient, and sweep Sent Items afterwards |
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

**Removed (#68).** The `llm-tests/` pytest harness was deleted along with the rest of the
PowerPoint surface: every scenario it contained targeted charts, tables, ranges, slides and
styling, none of which exist in this repository any more. `scripts/Test-LlmRegressionGate.ps1`
and the `run_llm_gate` workflow input went with it.

If LLM-behaviour testing is wanted again, it should be designed against the five Outlook tools
from scratch rather than resurrected from the PowerPoint scenarios.
