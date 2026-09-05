# ADR-001: Why OutlookMcp Has No Traditional Unit Tests

**Status**: Accepted (amended 2026-09-02)  
**Date**: 2025-11-02  
**Decision Makers**: Architecture Team  
**Stakeholders**: Development Team, Code Reviewers, Contributors

> **Amendment, 2026-09-02 (#37).** As written, this ADR and Rule 30 said "never write unit tests",
> while the repository contained 16 files under `tests/**/Unit/`. A rule contradicted by the codebase
> it governs is not enforceable, so #37 asked for an explicit decision between narrowing the rule,
> deleting all 16 files, or retiring the rule.
>
> **Decision: narrow the rule.** The ban is on *mocked-COM* unit tests, which is what the rationale
> below actually argues against. Tests over genuinely pure logic - enum and action-string mappings,
> JSON parsing, HRESULT classification, result-type invariants, guard clauses that return before any
> COM call - are permitted, because they exercise real code paths and would otherwise go unverified.
> The "Exceptions" section below is now the normative list; it replaces the earlier hypothetical
> phrasing ("if we had complex calculations independent of Outlook (we don't)"), which was simply
> untrue by the time it was read.
>
> Three files were deleted under this decision: `ComUtilitiesTests`, `ComUtilitiesExtendedTests`, and
> `PptContextTests`. All three claimed to test COM behaviour while passing `null!` or plain strings
> where a COM object belonged - one test was even named `Release_WithComObject_DoesNotThrow` above a
> comment admitting no COM object was involved. They are exactly what this ADR was written to prevent.
>
> One file is knowingly left in place: `OleMessageFilterTests`, which is **currently failing on
> `master`** (#59). Deleting it would erase a live signal before anyone has established whether the
> filter or the test is wrong, so its fate is deferred to #59.
>
> This amendment does not weaken the core position: **no test may stand in for an integration test.**
> A permitted pure test proves only its own narrow claim. It never demonstrates that a COM operation
> works, and it must not be cited as coverage for one.

---

## Context and Problem Statement

OutlookMcp is a COM automation library that wraps the classic Outlook for Windows COM API. During code review, the question inevitably arises: **"Why don't you have unit tests?"**

This ADR documents our architectural decision and the reasoning behind our testing strategy.

Only the **classic Outlook for Windows desktop app** can be automated. The new Outlook for Windows exposes no COM object model, and Office COM automation is unavailable on Linux and macOS.

---

## Decision

**We do NOT write traditional unit tests for OutlookMcp's COM-dependent behavior.** Our Outlook behavior coverage consists of **integration tests** that interact with a real classic Outlook instance via COM automation.

### What We DON'T Do

❌ Mock Outlook COM objects  
❌ Write mocked tests for mailbox, mail, calendar, folder, or attachment behavior  
❌ Test COM-dependent internal methods in isolation  
❌ Count pure .NET tests as coverage for Outlook automation  

### What We DO Do

✅ Write integration tests against real classic Outlook  
✅ Test Outlook operations through the actual COM object model  
✅ Verify behavior by re-reading real Outlook state  
✅ Keep pure tests only for logic that reaches no COM object at all  
✅ State honestly that hosted CI does not verify Outlook behavior today  

---

## Rationale

### 1. Outlook COM Cannot Be Meaningfully Mocked

**The Problem**: Outlook's COM API is the external system we're automating against. Every Outlook operation goes through `OutlookInteropRunner.Execute(...)`, which marshals work onto the single process-wide STA thread owned by `OutlookMcp.ComInterop.OutlookDispatcher`. Consider this simplified shape from the real folder commands:

```csharp
public OutlookFolderResolveResult ResolvePath(string? folder = null, bool includeItemCount = true)
{
    return OutlookInteropRunner.Execute(
        "OutlookFolderResolvePath",
        (application, session) =>
        {
            Outlook.Explorer? explorer = null;
            Outlook.MAPIFolder? resolvedFolder = null;
            object? items = null;

            try
            {
                resolvedFolder = OutlookInteropRunner.ResolveFolder(
                    application,
                    session,
                    folder,
                    DefaultFolderAliases,
                    ref explorer);

                if (resolvedFolder == null)
                {
                    return new OutlookFolderResolveResult
                    {
                        Success = false,
                        RequestedFolder = folder,
                        Resolved = false,
                        ErrorMessage = BuildUnknownFolderMessage(folder)
                    };
                }

                if (includeItemCount)
                {
                    items = resolvedFolder.Items;
                }

                return new OutlookFolderResolveResult
                {
                    Success = true,
                    RequestedFolder = folder,
                    Resolved = true,
                    Name = resolvedFolder.Name,
                    FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder)
                };
            }
            finally
            {
                OutlookInteropRunner.ReleaseComObject(ref items);
                OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                OutlookInteropRunner.ReleaseComObject(ref explorer);
            }
        },
        ex => new OutlookFolderResolveResult
        {
            Success = false,
            RequestedFolder = folder,
            Resolved = false,
            ErrorMessage = $"Failed to resolve the Outlook folder: {ex.Message}"
        });
}
```

**What would a "unit test" look like?**

```csharp
// Option 1: Mock the COM object
var mockNamespace = new Mock<dynamic>();  // ❌ Cannot meaningfully mock dynamic COM objects
mockNamespace.Setup(n => n.GetDefaultFolder(...)).Returns(...);  // ❌ Runtime binding lies

// Option 2: Test without Outlook
[Fact]
public void ResolvePath_ReturnsSuccess()
{
    var result = ResolvePath(null!);  // ❌ No Outlook.NameSpace, no MAPIFolder
    Assert.True(result.Success);      // ❌ This proves nothing
}
```

**The Truth**: The ONLY way to verify this code works is to:
1. Attach to a real running classic Outlook instance
2. Call the real COM API on the dispatcher STA thread
3. Verify the folder, mail item, appointment, or draft actually exists or changed in Outlook

**That's an integration test by definition.**

### 2. Our Integration Tests ARE Our Unit Tests

In traditional layered architecture:
- **Unit tests** test business logic in isolation
- **Integration tests** verify components work together
- **E2E tests** test the entire system

In COM automation architecture:
- **Integration tests** test business logic AND COM interaction (these ARE our unit tests)
- **E2E tests** don't exist (we ARE the library, not an application)

**Analogy**: OutlookMcp is like a database driver (e.g., Npgsql for PostgreSQL):
- You don't mock `DbConnection` to prove SQL semantics
- You test against a real database instance
- The "integration test" IS the unit test

### 3. Industry Precedent

This pattern is **normal and correct** for COM/browser/external system automation:

| Library | What It Automates | Test Strategy |
|---------|------------------|---------------|
| **Selenium WebDriver** | Browser DOM | Integration tests against real browsers |
| **Playwright** | Browser automation | Integration tests with browser instances |
| **AWS SDK** | Cloud services | Integration tests against AWS (or LocalStack) |
| **OutlookMcp** | Outlook COM | Integration tests against classic Outlook |

**None of these libraries have meaningful unit tests** for their core automation logic. They all test against the real external system.

### 4. What About .NET Framework APIs?

**Question**: "Shouldn't we unit test our wrappers around .NET APIs?"

**Answer**: No, because .NET already tests those APIs. Consider:

```csharp
public static string ValidateAndNormalizePath(string path)
{
    if (string.IsNullOrWhiteSpace(path))
        throw new ArgumentException("Path cannot be null");
    
    return Path.GetFullPath(path);  // .NET handles validation
}
```

**What's actually happening**:
- `Path.GetFullPath()` does: path traversal prevention, invalid character checking, normalization
- Our code does: null check (trivial)

**Testing this**:

```csharp
[Fact]
public void ValidatePath_WithTraversal_ThrowsException()
{
    Assert.Throws<ArgumentException>(() =>
        PathValidator.ValidateAndNormalizePath("../../etc/passwd"));
}
```

**Problem**: This test verifies .NET's `Path.GetFullPath()` works, not our code. We're testing Microsoft's code, not ours.

**Better approach**: Trust .NET's APIs (they're battle-tested). If our path validation is wrong, our integration tests will fail when we try to resolve a folder or file attachment path through Outlook.

### 5. The MCP Protocol Argument

**Question**: "Shouldn't we unit test MCP JSON serialization?"

**Answer**: No, the MCP SDK handles protocol compliance.

```csharp
public class MailListResult : ResultBase
{
    public List<MailItemInfo> Items { get; set; } = [];
}

// MCP SDK serializes this to JSON automatically
```

**What a unit test would look like**:

```csharp
[Fact]
public void MailListResult_SerializesToJson()
{
    var result = new MailListResult
    {
        Items = [new MailItemInfo { Subject = "Hello" }]
    };
    var json = JsonSerializer.Serialize(result);
    Assert.Contains("Hello", json);
}
```

**Problem**: This tests `System.Text.Json`, not our code. If JSON serialization breaks, the MCP SDK will fail to parse responses, and protocol integration tests will catch it.

---

## Real-World Test Coverage

### What Our Integration Tests Actually Test

**Scenario**: Resolve the Inbox and inspect its item payload.

```csharp
[Fact]
[Trait("Category", "Integration")]
[Trait("Feature", "OutlookSeed")]
public void Folder_ListItems_Inbox_ReturnsMailboxPayload()
{
    // Act - real Outlook COM via OutlookInteropRunner and OutlookDispatcher
    var resolveResult = _folderCommands.ResolvePath("inbox");
    var listResult = _folderCommands.ListItems("inbox", maxCount: 10, includePreview: false);

    // Assert - verify actual Outlook state, not just a wrapper return value
    Assert.True(resolveResult.Success);
    Assert.True(resolveResult.Resolved);
    Assert.False(string.IsNullOrWhiteSpace(resolveResult.FolderPath));

    Assert.True(listResult.Success);
    Assert.Equal("Inbox", listResult.FolderName, ignoreCase: true);
    Assert.True(listResult.ReturnedCount <= 10);
    Assert.All(listResult.Items, item =>
    {
        Assert.False(string.IsNullOrWhiteSpace(item.ItemType));
        Assert.False(string.IsNullOrWhiteSpace(item.MessageClass));
    });
}
```

**Scenario**: Create a draft, modify it, then re-read it through a fresh call.

```csharp
[Fact]
[Trait("Category", "Integration")]
[Trait("Feature", "OutlookSeed")]
public void Mail_CreateDraft_SetSubject_Read_VerifiesMailboxState()
{
    var createResult = _mailCommands.CreateMailDraft(
        subject: "OutlookMcp integration seed",
        body: "Created by an integration test.");
    Assert.True(createResult.Success);
    Assert.False(string.IsNullOrWhiteSpace(createResult.EntryId));

    var updateResult = _mailCommands.SetSubject(
        "OutlookMcp integration seed updated",
        entryId: createResult.EntryId,
        storeId: createResult.StoreId,
        useActiveMail: false);
    Assert.True(updateResult.Success);

    var readResult = _mailCommands.Read(
        entryId: createResult.EntryId,
        storeId: createResult.StoreId,
        useActiveMail: false);
    Assert.True(readResult.Success);
    Assert.Equal("OutlookMcp integration seed updated", readResult.Subject);
}
```

**What this ACTUALLY tests**:
1. ✅ Outlook process attachment (`Outlook.Application` via `GetActiveObject`)
2. ✅ Single process-wide STA dispatch (`OutlookDispatcher.Shared.Execute`)
3. ✅ COM object lifecycle (`Outlook.NameSpace`, `MAPIFolder`, `Items`, `MailItem`)
4. ✅ Outlook Object Model Guard error surfacing
5. ✅ Error handling (`Success` false with `ErrorMessage`)
6. ✅ Resource cleanup (`OutlookInteropRunner.ReleaseComObject(ref obj)`)
7. ✅ Re-reading mailbox state through a fresh command call
8. ✅ Business logic (folder resolution, draft creation, mail mutation)
9. ✅ API contract (`IFolderCommands`, `IMailCommands`)

**A unit test could verify**: None of the above (requires real classic Outlook).

### Test Statistics

- **Outlook command domains**: 8 (`Application`, `Attachment`, `Calendar`, `Contact`, `Folder`, `Mail`, `Rule`, `Task`)
- **Current product surface**: 66 operations across those domains
- **Hosted CI Outlook coverage**: 0 operations
- **Manual Outlook smoke coverage**: `Feature=OutlookSeed`
- **False Positives**: Lower when tests use real Outlook state instead of mocked COM

---

## Consequences

### Positive

✅ **Tests verify real behavior** - We test what actually happens in Outlook, not mocked abstractions  
✅ **High confidence** - A passing local Outlook integration test proves the operation works against the real COM surface  
✅ **No mock maintenance** - No complex mock setup that becomes outdated  
✅ **Catches integration bugs** - We discover COM quirks such as STA affinity, RCW lifetime, Object Model Guard prompts, and MAPI item shape differences  
✅ **Industry standard** - Follows proven patterns from Selenium, Playwright, AWS SDK  

### Negative

⚠️ **Slower tests** - Real Outlook tests are slower than pure .NET checks  
⚠️ **Requires classic Outlook** - Integration tests need Windows, classic Outlook installed, running, and signed in  
⚠️ **Shared application state** - Outlook is a long-running desktop application with a real mailbox, not a disposable test document  
⚠️ **Cannot run on Linux/macOS** - Office COM automation is Windows-only  
⚠️ **Hosted CI gap** - GitHub-hosted runners can build and run CI-safe tests, but cannot automate Outlook  

### Mitigation Strategies

**For slow tests**:
- Run the smallest targeted test command for the changed layer
- Prefer CI-safe metadata and protocol tests for fast feedback
- Use `Feature=OutlookSeed` for manual Outlook smoke coverage
- Keep destructive Outlook scenarios explicit and cleaned up

**For Outlook dependency**:
- Local Outlook integration testing requires classic Outlook for Windows, already running and signed in
- Run the test process at the same elevation level as Outlook; `GetActiveObject` cannot cross integrity levels
- Treat tests that touch Outlook as real mailbox automation
- There is no CI execution of Outlook tests and none is planned (#31, closed as not planned): local runs are the only verification

---

## Alternatives Considered

### Alternative 1: Mock Outlook COM Objects

**Rejected** because:
- `dynamic` COM objects cannot be meaningfully mocked
- Mocks would just verify our mock setup, not real Outlook behavior
- Outlook's COM API has quirks (single-instance process model, STA dispatch, Object Model Guard, MAPI folder/item shape) that mocks would not catch

### Alternative 2: Record/Replay COM Interactions

**Rejected** because:
- Fragile (breaks when Outlook updates or mailbox state changes)
- Doesn't test actual Outlook state
- High maintenance burden
- Doesn't verify the Object Model Guard, dispatcher, or live MAPI behavior

### Alternative 3: Separate Business Logic from COM

**Rejected** because:
- There IS no business logic separate from COM interaction for Outlook behavior
- Our "business logic" IS calling Outlook COM methods correctly
- Would create artificial abstraction layers with no value

### Alternative 4: Test Against Outlook Interop Primary Assemblies

**Rejected** because:
- Still requires classic Outlook installed
- PIAs are just type definitions, not implementation
- Doesn't reduce test execution time
- We use COM objects at runtime and must verify real runtime behavior

---

## Exceptions: When Unit Tests Make Sense

**This section is normative** (see the 2026-09-02 amendment at the top). A unit test is permitted
only if it satisfies **all** of the following:

1. It touches **no COM object at all** - not a real one, not a `null!` stand-in, not a mock.
2. Its subject is genuinely pure: a mapping, a parse, a classification, an invariant, or a guard
   clause that returns before any COM call is reached.
3. It would fail if the logic under test were wrong. (If it can only fail when .NET itself is
   broken, it is testing Microsoft's code, not ours - see sections 4 and 5 above.)
4. It carries `[Trait("Category", "Unit")]` and lives under `tests/**/Unit/`, so it is trivially
   separable from the integration suite.

If a test needs a COM object to mean anything, and you are substituting something for that object
to make it run, the test is prohibited. Write an integration test instead, or write none.

### Currently permitted under this exception

| Test file | Why it qualifies |
|---|---|
| `ActionValidatorTests` | Enum-to-action-string mapping over generated metadata (Rule 15 guard) |
| `CoreCommandsCoverageTests` | Reflection over `[ServiceCategory]`; guards MCP/CLI surface coverage |
| `ResultTypeInvariantTests` | Enforces the Rule 1 `Success`/`ErrorMessage` invariant by reflection |
| `ResultTypeSerializationTests` | JSON shape of result contracts |
| `ServiceRegistryJsonParsingTests` | Parsing of generated registry JSON |
| `ParameterTransformsFileTests` | Pure parameter transformation |
| `McpServerVersionCheckerTests` | Version comparison logic |
| `OutlookInteropRunnerTests` | HRESULT classification for Object Model Guard denials (#30) |
| `MailCommandsSendTests` | Send confirmation gate and idempotency cache; returns before COM (#29) |
| `DestructiveConfirmationGateTests` | Confirmation gates that are guard clauses over the caller's arguments and return before COM (#9) |
| `OutlookDispatcherTests` | STA queue/serialization mechanics with plain delegates (#20) |
| `StreamJsonRpcTests` | Real in-process duplex streams; no COM in the RPC layer |
| `OutlookMcpServiceErrorTests` | Error-message formatting regression guard |
| `ConfigurationReloadTests` | `reloadOnChange` configuration regression guard |

Two entries deserve their caveats stated rather than buried:

- **`OutlookDispatcherTests`** is the weakest entry here. It proves the dispatcher serializes work
  onto one STA thread, which is real and worth guarding, but it cannot prove that Outlook COM calls
  made *through* that thread behave correctly. It is a placeholder for the real regression test that
  #20 asked for. That real regression test is no longer blocked on CI - #31 was closed as not
  planned, and the Outlook-backed regression test now exists and runs locally
  (`Execute_AfterPriorCall_SharedApplicationRemainsUsable`). What remains missing is *enforcement*,
  not coverage: nothing makes a contributor run it.
- **`OleMessageFilterTests`** is *not* in the table. It tests COM plumbing without COM, so it does
  not qualify - but it is currently failing on `master`, and deleting a failing test is how real
  defects get lost. Resolution is deferred to #59.

**The standing expectation is unchanged**: nearly all logic in this repository involves COM
interaction, so nearly all Outlook behavior tests must be integration tests.

---

## Code Review Response Template

When reviewers ask "Why no unit tests?", respond:

> **OutlookMcp is a COM automation library.** We test against real classic Outlook because:
> 
> 1. **Outlook COM cannot be mocked** - Dynamic COM objects don't support meaningful traditional mocking
> 2. **Integration tests ARE our unit tests** - We test COM interaction in the only way possible
> 3. **Industry standard** - Selenium, Playwright, AWS SDK all use the same pattern
> 4. **High confidence** - Tests verify actual Outlook behavior, not mock abstractions
> 
> See `docs/ADR-001-NO-UNIT-TESTS.md` for full rationale.

---

## References

1. **Martin Fowler - "Test Pyramid Antipattern"**: https://martinfowler.com/bliki/TestPyramid.html
   - "The test pyramid is a simplification... some contexts don't fit the pyramid"
   
2. **Selenium Testing Best Practices**: https://www.selenium.dev/documentation/test_practices/
   - Tests run against real browsers, not mocks
   
3. **Playwright Testing Philosophy**: https://playwright.dev/docs/test-philosophy
   - "End-to-end tests should test real scenarios"
   
4. **AWS SDK Testing**: https://github.com/aws/aws-sdk-net
   - Integration tests against AWS or LocalStack, minimal unit tests

5. **Microsoft Office Interop Best Practices**: https://learn.microsoft.com/office/client-developer/
   - COM automation testing requires real Office instances

---

## Decision Record

**Date**: November 2, 2025  
**Decided by**: Architecture Team  
**Status**: Accepted  

**Supersedes**: N/A  
**Superseded by**: N/A  

**Last Reviewed**: September 2, 2026  
**Next Review**: When adding Outlook features that do not require COM (if ever)

---

## Appendix: Test Execution Strategy

### Local Development

```powershell
# Everything that runs without Outlook
dotnet test --filter "Category!=Integration"
```

```powershell
# Targeted examples by current Feature trait
dotnet test tests\OutlookMcp.McpServer.Tests\OutlookMcp.McpServer.Tests.csproj --filter "Feature=McpProtocol"
dotnet test tests\OutlookMcp.ComInterop.Tests\OutlookMcp.ComInterop.Tests.csproj --filter "Feature=OutlookDispatcher"
dotnet test tests\OutlookMcp.CLI.Tests\OutlookMcp.CLI.Tests.csproj --filter "Feature=CliExitCode"
```

```powershell
# Manual Outlook smoke tests on Windows with classic Outlook running
dotnet test tests\OutlookMcp.Core.Tests\OutlookMcp.Core.Tests.csproj --filter "Feature=OutlookSeed"
```

### Pre-Commit

```powershell
# Fast CI-safe guard for generated Core/MCP surface coverage
dotnet test tests\OutlookMcp.McpServer.Tests\OutlookMcp.McpServer.Tests.csproj --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

### CI/CD Pipeline

- **GitHub-hosted runners**: build verification and CI-safe tests only; they do not verify Outlook COM behavior.
- **`.github\workflows\integration-tests.yml`**: wired but disabled. It reports `integration-runner-disabled` while `ENABLE_OUTLOOK_INTEGRATION_CI` is not `true`.
- **Self-hosted Outlook runner**: would be required before CI could run Outlook integration tests - a Windows runner labelled `outlook` with classic Outlook installed, running, and signed in. **There is no plan to provide one** (#31, closed as not planned).
- **Merge posture**: never claim automated Outlook runtime verification. It does not exist and is not coming. Outlook claims rest on a local run against a real profile, or they are unverified and should be labelled as such.

---

**End of ADR-001**
