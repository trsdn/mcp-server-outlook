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
`OutlookDispatcher`, `OutlookMcpService`, `OutlookSeed`, `ParameterTransforms`, `RuleCrud`,
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

### Assert a property of the code, not of the mailbox

The subtler version of the same failure. The assertion is binary, the test executes, it passes —
and it is testing the fixture rather than the behaviour.

```csharp
// WRONG: asserts a property of the data. True of most received mail, not all.
Assert.Contains(result.Headers, h => h.Name == "Received");
```

That one passed review and passed on the mailbox, then failed later when the first inbox message
carrying transport headers turned out to be an internal notification with twenty headers and no
`Received` line. Nothing about the parser had changed.

Three ways this shows up here, all found in real work:

- **Naming data you did not create.** If the test needs a specific header, category, sender or
  folder, either create it or derive it from what the mailbox actually returned. Do not name it.
- **A fixture that quietly sidesteps the logic.** A paging test built on freshly created drafts
  gets *distinct* `ReceivedTime` values, so it never enters the tied-timestamp band it was written
  to cover, and passes having exercised none of it. Ask what the fixture would have to look like
  for the test to be meaningful, and assert that it does.
- **Reusing the implementation to check the implementation.** A test that parses with the parser's
  own helper proves only that the parser agrees with itself.

The first two are the same failure from opposite directions: one over-assumes the data, the other
over-controls it. In both, the fixture is what is really being asserted about, and the code under
test is never pushed into the state the test exists to cover.

The technique that works: **verify by an independent, deliberately simpler route.** Rather than
asserting which headers came back, count the header-start lines in the raw block with a second
naive implementation written in the test, and assert the parsed count matches. That assumes nothing
about which headers exist, and still fails if the block comes back unsplit or if continuation lines
leak through. Where a filter is involved, assert the filtered count equals the count of that name in
the unfiltered set — which catches dropped duplicates as well as extras.

**"Simpler" has to mean strictly weaker, or it is not a cross-check.** Reimplementing the parser in
the test would reproduce the same folding bug in both places and agree with itself just as happily
as calling the parser's own helper would. Counting header-start lines works precisely *because* it
establishes a count and nothing about the values. If a check like this ever needs to grow toward
the shape of the thing it is checking, it has stopped being independent.

And check the numbers are non-trivial. "20 headers from 1767 characters" is evidence the test did
something; `0 == 0` is not.

---

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Only asserting the success flag | Verify actual Outlook state |
| "Accept both" assertions | Binary assertions only |
| Asserting a property of the mailbox's data | Create the data, or derive the expectation from what was returned |
| Verifying a parser with the parser's own helper | Check by an independent, simpler route |
| Mocking `Outlook.Application` or `NameSpace` | See ADR-001 - mocked COM proves nothing |
| Mutating the user's real mailbox in a test | Operate on drafts you created, and clean up |
| Missing `Feature` trait | Add one from the list above |
| Leaving an integration test undeclared | Mark `[Trait("Category", "Integration")]` |

---

## A red test on a sending surface really does send

Rule 29 requires watching the test fail before implementing. On `send`, `forward`, `reply-all` or
anything that books a meeting, "watch it fail" means the guard is not there yet and **the action
happens** — from the user's real mailbox. That is inherent to TDD here, not carelessness, but it
has to be planned for rather than discovered.

Address anything a red test might send to the reserved `.invalid` TLD, which can never be
delivered. Expect it to reach Sent Items anyway, and clean up afterwards. This has already happened
once during work on #9; the message was undeliverable and was removed, but only because it had been
addressed defensively in advance.

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
