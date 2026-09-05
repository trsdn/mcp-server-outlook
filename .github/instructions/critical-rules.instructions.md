---
applyTo: "**"
---

# CRITICAL RULES - MUST FOLLOW

> **⚠️ NON-NEGOTIABLE rules for all OutlookMcp development**

## Rule 0: NEVER Commit Without Running Tests (CRITICAL)

**NEVER commit, push, or create PRs without first running tests for the code you changed.**

**Why Critical:** Prevents breaking changes from reaching main, wastes team time debugging failures, violates CI/CD principles.

**Enforcement:**
- ALWAYS run relevant tests BEFORE committing
- Use `--filter "Feature=<feature-name>&RunType!=OnDemand"` for feature-specific tests
- Build must pass (0 warnings, 0 errors)
- Pre-commit hooks must pass (COM leaks, success flag, coverage)
- Document test results in commit message

**Process:**
1. Make code changes
2. Build: `dotnet build` (must succeed with 0 warnings)
3. Run tests: `dotnet test --filter "Feature=<feature>&RunType!=OnDemand"`
4. Verify all tests pass
5. Run pre-commit checks
6. THEN commit

## Rule 21: Never Commit Automatically (CRITICAL)

**NEVER commit or push code automatically. All commits, pushes, and merges must require explicit user approval.**

**Why Critical:** Prevents accidental changes, enforces review, and ensures user control over all repository modifications.

**Enforcement:**
- All automated tools, scripts, and agents must prompt for user approval before any commit, push, or merge.
- No background or silent commits allowed.
- Document this rule in all agent and automation instructions.

## Rule 26: No Confidential Information in Commits/PRs (CRITICAL)

**NEVER include confidential project names, file names, customer names, or internal references in commit messages, PR descriptions, or issue descriptions.**

**Why Critical:** Git history and GitHub issues/PRs are public or semi-public. Confidential information in commits persists forever in git history and is difficult to remove.

**Forbidden in commits/PRs/issues:**
- Customer project names (e.g., "CP Toolkit", "Contoso Deal")
- Specific file paths or mailbox names from customer environments
- Internal tool names that reveal customer context
- Any information that could identify a specific customer engagement

**Allowed:**
- Generic descriptions ("a mail item", "a shared mailbox")
- Technical details that don't identify the source ("a column with a hyphen in the name")
- Error messages and stack traces (sanitized of paths/names)

**Example:**
```
# ❌ WRONG: Reveals confidential project
Discovered while debugging the Q4 Migration mailbox for Contoso

# ✅ CORRECT: Generic description
Discovered while debugging a mail item whose sender address could not be resolved
```

**Enforcement:**
- Review all commit messages before pushing
- Review all PR/issue descriptions before submitting
- If confidential info is committed, immediately amend or rebase to remove it


## Quick Reference (Grouped by Context)

**Note:** Rules below are grouped by when you need them, not by number. Detailed rules follow in numeric order (1-21).

**Every Edit:**
| Rule | Action | Why Critical |
|------|--------|--------------|
| 0. Test before commit | ALWAYS run tests before committing | Prevents breaking changes |
| 1. Success flag | NEVER `Success=true` with `ErrorMessage` | Confuses LLMs, causes silent failures |
| 1b. No exception wrapping | Never catch in the action lambda; use the runner's onException | Prevents double-wrapping, preserves the runner's error classification |
| 16. Test scope | Only run tests for code you changed | Saves 10+ minutes per test run |
| 8. TODO markers | Must resolve before commit | Pre-commit hook blocks |
| 23. IDE warnings | TRUST them - never dismiss as false positives without verification | Prevents broken code |

**When Writing Code:**
| Rule | Action | Why Critical |
|------|--------|--------------|
| 29. TDD | Write test FIRST → RED → implement → GREEN | Proves tests catch real bugs |
| 30. Integration tests | No mocked-COM unit tests; narrow pure-logic exception only | Unit tests prove nothing for COM interop |
| 22. COM cleanup | ALWAYS use try-finally, NEVER swallow exceptions | Prevents leaks and silent failures |
| 7. COM API | Use Outlook COM first, validate docs | Prevents wrong dependencies |
| 9. GitHub search | Search OTHER repos for Outlook VBA/COM examples FIRST | Learn from working code |
| 2. NotImplementedException | Never use, full implementation only | No placeholders allowed |
| 15. Enum mappings | All enum values mapped in ToActionString() | Runtime errors otherwise |
| 17. MCP error checks | Always return JSON, check result.Success | MCP protocol requirement |

**When Writing Tests:**
| Rule | Action | Why Critical |
|------|--------|--------------|
| 2. Tests | Fail loudly, never silent | Silent failures waste hours |
| 30. Integration tests | No mocked-COM unit tests; narrow pure-logic exception only | Unit tests prove nothing for COM interop |
| 15. No Save | Remove unless testing persistence | Makes tests 50% faster |
| 11. Test debugging | Run tests one by one | Isolates actual failure |
| 13. Test compliance | Pass checklist before PR submission | Prevents test pollution |

**Before Commit:**
| Rule | Action | Time |
|------|--------|------|
| 0. Test before commit | ALWAYS run tests before committing | 3-5 min |
| 4. Session code | See testing-strategy for test commands | 3-5 min |
| 6. COM leaks | Pre-commit hook runs it - but see Rule 5, a clean run is not evidence | 1 min |
| 7. PRs | Always use PRs, never direct commit | Always |
| 24. Post-change sync | Verify ALL sync points (CLI, SKILLs, READMEs, counts) | 5-10 min |
| 26. No confidential info | No customer/project names in commits/PRs/issues | Always |
| 27. CHANGELOG | Update CHANGELOG.md before merging PRs | 2-5 min |
| 28. COM API naming | Match COM param names when clear in flat schema | Always |

**During PR Process:**
| Rule | Action | Time |
|------|--------|------|
| 20. PR review comments | Check/fix automated comments immediately | 5-10 min |
| 14. Bug fixes | Complete 6-step process | 30-60 min |
| 19. Tool descriptions | Verify XML docs (`/// <summary>`) match behavior | Per change |

**Rare/Specialized:**
| Rule | Action | When |
|------|--------|------|
| 12. No test refs | Production NEVER references tests | Architecture review |
| 5. Instructions | Update after significant work | After major features |

---

## Rule 1: Success Flag Must Match Reality (CRITICAL)

**NEVER set `Success = true` when `ErrorMessage` is set. This is EXTREMELY serious!**

```csharp
// ❌ CRITICAL BUG: Confuses LLMs and users
result.Success = true;
result.ErrorMessage = "Draft created but recipients could not be resolved...";

// ✅ CORRECT: Success only when NO errors
if (!loadResult.Success) {
    result.Success = false;  // MUST be false!
    result.ErrorMessage = $"Failed: {loadResult.ErrorMessage}";
}
```

**Invariant:** `Success == true` ⟹ `ErrorMessage == null || ErrorMessage == ""`

**Why Critical:** LLMs see Success=true and assume operation worked, causing workflow failures and silent data corruption.

**Common Bug Pattern (43 violations found 2025-01-28):**
```csharp
// ❌ WRONG: Optimistic Success setting without catch block correction
var result = new OperationResult();
result.Success = true;  // Set optimistically

try {
    // ... do work ...
    return result;
} catch (Exception ex) {
    // ❌ BUG: Forgot to set Success = false!
    result.ErrorMessage = $"Error: {ex.Message}";
    return result;  // Returns Success=true with ErrorMessage! 
}

// ✅ CORRECT: Set Success in try block, always false in catch
var result = new OperationResult();

try {
    // ... do work ...
    result.Success = true;  // Only set true on actual success
    return result;
} catch (Exception ex) {
    result.Success = false;  // ✅ Always false in catch!
    result.ErrorMessage = $"Error: {ex.Message}";
    return result;
}
```

**Enforcement:**
- Pre-commit hook runs `check-success-flag.ps1` to detect violations
- Regression tests verify this invariant (SlideSuccessErrorRegressionTests)
- Code review MUST check every `Success = ` assignment
- Search pattern: `Success.*true.*ErrorMessage`

**Examples of bugs found:**
- 43 violations were found and fixed across the Core command surface
- All followed pattern: `Success = true` at start, `ErrorMessage` set in catch without `Success = false`

---

## Rule 1b: Never Suppress Exceptions with Try-Catch (CRITICAL)

**Core Commands: NEVER add a catch inside the action lambda that returns an error result. Let exceptions reach the runner's `onException` delegate.**

```csharp
// CRITICAL BUG: suppressing the exception with a local error result
public OperationResult CreateDraft(string subject)
{
    return OutlookInteropRunner.Execute(
        "OutlookMailCreateDraft",
        (application, session) =>
        {
            try
            {
                Outlook.MailItem mail = (Outlook.MailItem)application.CreateItem(
                    Outlook.OlItemType.olMailItem);
                mail.Subject = subject;
                mail.Save();
                return new OperationResult { Success = true };
            }
            catch (Exception ex)
            {
                // WRONG: replaces the runner's classified error with a generic one
                return new OperationResult { Success = false, ErrorMessage = ex.Message };
            }
        },
        ex => new OperationResult { Success = false, ErrorMessage = ex.Message });
}

// CORRECT: let it propagate to onException
public OperationResult CreateDraft(string subject)
{
    return OutlookInteropRunner.Execute(
        "OutlookMailCreateDraft",
        (application, session) =>
        {
            Outlook.MailItem? mail = null;
            try
            {
                mail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);
                mail.Subject = subject;
                mail.Save();
                return new OperationResult { Success = true };
            }
            finally
            {
                // Finally for cleanup, not catch for error handling
                OutlookInteropRunner.ReleaseComObject(ref mail);
            }
        },
        ex => new OperationResult
        {
            Success = false,
            ErrorMessage = $"Failed to create the draft: {ex.Message}"
        });
}
```

**Why Critical:**
- `OutlookInteropRunner.Execute` already classifies COM failures, Object Model Guard denials,
  "Outlook not running", and elevation mismatches into actionable messages
- A local catch replaces those messages with a generic one and originates from the wrong layer
- `onException` is the single, predictable place a failure becomes a result
- Finally blocks are CORRECT for resource cleanup; catch blocks are not for error suppression

**Safe Patterns (Keep these):**
- Loop continuations: `catch { continue; }` when one inaccessible item should not fail a listing
- Optional property access: `catch { propValue = null; }` where absence is meaningful and handled
- Specific error routing: `catch (COMException ex) when (ex.HResult == code) { handle... }`
- Finally blocks: resource cleanup for COM objects

**Pattern to Remove:**
- `catch (Exception ex) { return new Result { Success = false, ErrorMessage = ex.Message }; }`
  inside the action lambda

**Architecture Foundation:**
```
Core Command Method
  --> OutlookInteropRunner.Execute(operationName, action, onException)
      --> OutlookDispatcher.Shared marshals onto the STA thread
          --> action throws
              --> runner classifies (COM / OMG / not-running / elevation)
                  --> onException returns Result { Success = false, ErrorMessage }
```

**See:** architecture-patterns.instructions.md for the complete exception propagation pattern.

---

## Rule 2: No NotImplementedException

**Every feature must be fully implemented with real Outlook COM operations and passing tests. No placeholders.**

---

## Rule 3: Session Cleanup Tests

**When modifying session/batch code, run OnDemand tests (see testing-strategy.instructions.md). Must pass before commit.**

---

## Rule 4: Update Instructions

**After significant work, update `.github/copilot-instructions.md` with lessons learned, architecture changes, and testing insights.**

---

## Rule 5: COM Object Leak Detection

**Every COM object must be released in a `finally` block** using
`OutlookInteropRunner.ReleaseComObject(ref obj)` — except anything Outlook hands back from a cache
or that you reached *upward* via `.Parent` / `.Session`, which needs `ReleaseSharedComObject`
(see #19, #116, #122, #134 and the release-ownership section in
[outlook-com-interop](outlook-com-interop.instructions.md)).

Run `& "scripts\check-com-leaks.ps1"` before commit. **Do not treat a clean run as evidence that
you got this right.** It is a smoke test for the file that forgot cleanup entirely — useful for
that, and no evidence at all about the file that mostly remembered. Specifically (#136):

- it only examines files declaring an explicitly-typed `Outlook.X y =` local, roughly 11 of 74
  source files, so anything acquiring COM through `var` is invisible to it
- its verdict is two file-wide booleans with no pairing of acquisition to release, so one release
  call anywhere in a nine-hundred-line file marks every acquisition in it clean — it cannot catch a
  conditional-path leak by construction
- its release regex matches `ReleaseComObject` and `ReleaseSharedComObject` **identically**, so it
  is silent on the exact distinction above

All four use-after-free bugs in this codebase were in files it reported clean. The rule is the
`finally` discipline; the script is a convenience, not the check.

---

## Rule 6: All Changes Via Pull Requests

**Never commit to `main`. Create feature branch → PR → CI/CD + review → merge.**

**Enforcement:** Pre-commit hook blocks commits to main. If you're on main:
```bash
git stash                                    # Save changes
git checkout -b feature/your-feature-name    # Create feature branch
git stash pop                                # Restore changes
git add <files>                              # Stage changes
git commit -m "your message"                 # Commit to feature branch
```

**Why Critical:** Direct commits to main bypass CI/CD, skip code review, and violate branch protection.

---

## Rule 7: COM API First

**Use the Outlook COM API for everything it supports. Only reach for an external library for something Outlook COM genuinely cannot do.**

Validate against [the Outlook VBA reference](https://learn.microsoft.com/office/vba/api/overview/outlook) before adding dependencies.

---

## Rule 8: No TODO/FIXME Markers

**Code must be complete before commit. No TODO, FIXME, HACK, or XXX markers in source code.**

Delete commented-out code (use git history). Exception: Documentation files only.

---

## Rule 9: Search External GitHub Repositories for Working Examples First

**BEFORE creating new Outlook COM Interop code or troubleshooting COM issues:**

- **ALWAYS** search OTHER open source GitHub repositories for working examples
- **NEVER** search your own repository - only search external projects
- **NetOffice is THE BEST source for ALL COM Interop work**: https://github.com/NetOfficeFw/NetOffice
  - Strongly-typed C# wrappers for ALL Office COM APIs (Outlook, Excel, Word, etc.)
  - Search for ANY Outlook COM operation: MAPIFolder, Items.Restrict, MailItem, AppointmentItem, Attachments, Recipients, etc.
  - Study their patterns for dynamic interop conversion and proper COM object handling
  - NetOffice source is essentially a comprehensive reference for every Outlook COM API
- Look for repositories with Outlook automation, VBA code, or Office interop projects
- Search for the specific COM object/method you need (e.g., "Items.Restrict VBA date", "MailItem.Attachments VBA", "NameSpace.GetDefaultFolder NetOffice")
- Study proven patterns from other projects before writing new code
- Avoid reinventing solutions - learn from working implementations in the wild

**Why:** Outlook COM is quirky. Real-world VBA examples from other projects prevent common pitfalls (1-based indexing, the DASL/Jet dialect used by `Restrict`, object cleanup, the Object Model Guard, variant types).

---

## Rule 10: Debug Tests One by One

**When debugging test failures, ALWAYS run tests individually - never run all tests at once.**

**Process:**
1. List all test methods in the file
2. Run each test individually using `--filter "FullyQualifiedName=Namespace.Class.Method"`
3. Identify exact failure for each test before moving to next
4. Fix issues one test at a time

**Why:** Running all tests together masks which specific test fails and why. Individual execution provides clear, isolated diagnostics.

---

## Rule 11: Production Code NEVER References Tests

**Production code (Core, CLI, MCP Server) must NEVER reference test projects or test helpers.**

**Violations:**
- ❌ `<InternalsVisibleTo Include="*.Tests" />` in production `.csproj`
- ❌ `using OutlookMcp.*.Tests` in production code
- ❌ Production code calling test helper methods
- ❌ Production business logic in helper classes that tests use

**Correct Architecture:**
- ✅ **COM utilities** → `src/OutlookMcp.ComInterop/` (low-level COM helpers, dispatcher, message filter)
- ✅ **Business logic** → Private methods inside production Commands classes
- ✅ **Test helpers** → Call production commands, never duplicate logic
- ✅ `InternalsVisibleTo` only for production-to-production (e.g., Core → MCP Server)

**Why:** Tests depend on production code, not the reverse. Production code with test dependencies is broken architecture.

---

## Rule 12: Test Class Compliance Checklist

**Every new test class MUST pass the compliance checklist before PR submission.**

**Verify:**
- ✅ Uses `IClassFixture<TempDirectoryFixture>` (NOT manual IDisposable)
- ✅ Each test creates unique file via `CoreTestHelper.CreateUniqueTestFile()`
- ✅ NEVER shares test files between tests
- ✅ Tests that mutate the mailbox operate on drafts they created, and clean up
- ✅ Binary assertions only (NO "accept both" patterns)
- ✅ Required traits present (`Category`, `Feature`)
- ✅ Calls the Core command directly; no mocked COM interfaces
- ✅ NO duplicate helper methods (use CoreTestHelper)

**Why:** Systematic compliance prevents test pollution, file lock issues, silent failures, and maintenance nightmares. See [testing-strategy.instructions.md](testing-strategy.instructions.md) for complete checklist.

**Enforcement:** PR reviewers MUST check compliance before approval.

---

## Rule 13: Comprehensive Bug Fixes

**Every bug fix MUST include all 6 components before PR submission.**

**Required Components:**
1. ✅ **Code Fix** - Minimal surgical changes to fix root cause
2. ✅ **Tests** - Minimum 5-8 new tests (regression + edge cases + backwards compat)
3. ✅ **Documentation** - Update 3+ files (tool docs, user docs, prompts)
4. ✅ **Workflow Hints** - Update SuggestedNextActions and error messages
5. ✅ **Quality Verification** - Build passes, all tests green, 0 warnings
6. ✅ **PR Description** - Comprehensive summary (bug report, fix, tests, docs updated)

**Process:** Follow [bug-fixing-checklist.instructions.md](bug-fixing-checklist.instructions.md) for complete 6-step process.

**Why:** Incomplete bug fixes lead to regressions, confusion, and wasted time. Comprehensive fixes prevent future issues.

**Example:** the CLI exit-code bug (#63) = 1 code file + 9 tests + 3 doc files + a detailed PR description = complete fix.

---

## Rule 14: No Save Unless Testing Persistence

**Code must NOT call `batch.Save()` unless explicitly testing persistence.**

**Quick Rules:**
- ❌ FORBIDDEN: Tests only verifying operation success or in-memory state
- ✅ REQUIRED: round-trip tests verifying the change is visible on a subsequent read
- ⚡ REASON: Save is slow (~2-5s). Removing unnecessary saves makes tests 50%+ faster

**See:** [testing-strategy.instructions.md](testing-strategy.instructions.md) for complete Save patterns, when to use, and detailed examples.

---

## Quick Reference

| Rule | Action | Time |
|------|--------|------|
| 0. Test before commit | ALWAYS run tests before committing | 3-5 min |
| 1. Tests | Fail loudly, never silent | Always |
| 2. NotImplementedException | Never use, full implementation only | Always |
| 3. Session code | See testing-strategy for test commands | 3-5 min |
| 4. Instructions | Update after significant work | 5-10 min |
| 5. COM leaks | `finally` discipline per Rule 5; the script is a smoke test, not the check | 1 min |
| 6. PRs | Always use PRs, never direct commit | Always |
| 7. COM API | Use Outlook COM first, validate docs | Always |
| 8. TODO markers | Must resolve before commit | 1 min |
| 9. GitHub search | Search OTHER repos for Outlook VBA/COM examples FIRST | 1-2 min |
| 10. Test debugging | Run tests one by one, never all together | Per test |
| 11. No test refs | Production NEVER references tests | Always |
| 12. Test compliance | Pass checklist before PR submission | 2-3 min |
| 13. Bug fixes | Complete 6-step process (fix, test, doc, hints, verify, summarize) | 30-60 min |
| 14. No Save | See testing-strategy for complete patterns | Per test |
| 15. Enum mappings | All enum values mapped in ToActionString() | Always |
| 16. Test scope | Only run tests for code you changed | Per change |
| 17. MCP error checks | Check result.Success before JsonSerializer.Serialize | Every method |
| 18. Tool descriptions | Verify XML docs (`/// <summary>`) match tool behavior | Per tool change |
| 19. PR review comments | Check and fix all automated review comments after creating PR | 5-10 min |
| 24. Post-change sync | Verify ALL sync points (CLI, SKILLs, READMEs, counts) before commit | 5-10 min |
| 28. COM API naming | Match COM param names when clear in flat schema | Always |
| 29. TDD | Write test FIRST → RED → implement → GREEN | Always |
| 30. Integration tests | No mocked-COM unit tests; narrow pure-logic exception only | Always |



---

## Rule 15: Complete Enum Mappings (CRITICAL)

**Every enum value MUST have a mapping in ToActionString(). Missing mappings cause unhandled exceptions.**

```csharp
// ❌ WRONG: Incomplete mapping
public static string ToActionString(this RangeAction action) => action switch
{
    RangeAction.GetValues => "get-values",
    RangeAction.SetValues => "set-values",
    // Missing GetUsedRange, GetCurrentRegion, etc. → ArgumentException!
    _ => throw new ArgumentException($"Unknown RangeAction: {action}")
};

// ✅ CORRECT: All enum values mapped
public static string ToActionString(this RangeAction action) => action switch
{
    RangeAction.GetValues => "get-values",
    RangeAction.SetValues => "set-values",
    RangeAction.GetUsedRange => "get-used-range",  // ✅ All values
    RangeAction.GetCurrentRegion => "get-current-region",
    // ... all other values
    _ => throw new ArgumentException($"Unknown RangeAction: {action}")
};
```

**Why Critical:** Missing mappings cause MCP Server to throw exceptions instead of returning JSON, confusing LLMs.

**Enforcement:**
- Regression tests for all enum mappings
- When adding enum value, add mapping immediately
- Code review MUST verify completeness

**Example Bug:** `GetUsedRange` missing → "An error occurred invoking 'range'" (not JSON!)

---

## Rule 16: Test Only What You Changed (CRITICAL - PERFORMANCE)

**ALWAYS run tests ONLY for the specific code you modified. Integration tests take a very long time.**

**Wrong:**
```bash
# ❌ NEVER: Runs ALL integration tests (10+ minutes)
dotnet test --filter "Category=Integration&RunType!=OnDemand"
```

**Correct:**
```bash
# ✅ CORRECT: Test only the feature you changed
dotnet test --filter "Feature=OutlookDispatcher"   # Dispatcher / COM plumbing changes
dotnet test --filter "Feature=McpProtocol"         # MCP surface changes
dotnet test --filter "Feature=CliExitCode"         # CLI behaviour changes
```

**Why Critical:** integration tests require live Outlook COM automation and are slow. They also currently do not run in CI at all (#31), so a local run is the only signal you get.

**Enforcement:**
- Only run tests for files you modified
- Use Feature trait to target specific test groups
- Full test suite runs in CI/CD pipeline only

---

## Rule 17: MCP Tools Must Return JSON Responses (CORRECTED)

**Every MCP tool method that calls Core Commands MUST return JSON responses, not throw exceptions for business errors.**

```csharp
// ❌ WRONG: Throws exception for business logic errors
private static async Task<string> SomeAction(...)
{
    var result = await commands.SomeAsync(batch, param);
    
    if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
    {
        throw new ModelContextProtocol.McpException($"action failed: {result.ErrorMessage}");  // ❌ Wrong!
    }
    
    return JsonSerializer.Serialize(result, JsonOptions);
}

// ✅ CORRECT: Always return JSON - let result.Success indicate errors
private static async Task<string> SomeAction(...)
{
    var result = await commands.SomeAsync(batch, param);
    
    // Always return JSON (success or failure) - MCP clients handle the success flag
    return JsonSerializer.Serialize(result, JsonOptions);
}
```

**When to Throw McpException:**
- ✅ **Parameter validation** - missing required params, invalid formats
- ✅ **Pre-conditions** - file not found, batch not found, invalid state
- ❌ **NOT for business logic errors** - table not found, query failed, etc.

**Why:**
- ✅ MCP clients expect JSON responses with `success: false` for business errors
- ✅ HTTP 200 + JSON error = client can parse and handle gracefully
- ❌ HTTP 500 + exception = harder for clients to handle programmatically
- ✅ Core Commands return result objects with `Success` flag - serialize them!

**Example - Business Error (return JSON):**
```csharp
// Core returns: { Success = false, ErrorMessage = "Folder 'Archive' not found" }
// MCP Tool: Return this as-is
return JsonSerializer.Serialize(result, JsonOptions);
// Client gets: {"success": false, "errorMessage": "Folder 'Archive' not found"}
```

**Example - Validation Error (throw exception):**
```csharp
// Missing required parameter
if (string.IsNullOrWhiteSpace(tableName))
{
    throw new ModelContextProtocol.McpException("tableName is required for create-from-table action");
}
```

**Historical Note:** This rule was corrected on 2025-01-03 after discovering that tests expected JSON responses, not exceptions. The previous pattern (throwing McpException for business errors) was incorrect and caused MCP clients to receive unhandled errors instead of parseable JSON.

---

## Rule 18: Tool Descriptions Must Match Behavior (CRITICAL)

**Tool XML documentation (`/// <summary>`) is extracted by MCP SDK and sent to LLMs. It must be accurate and current.**

**What to verify when changing a tool:**

1. **Purpose and Use Cases Clear**:
   ```csharp
   // ❌ WRONG: Vague description
   /// <summary>Manage slides</summary>
   
   // ✅ CORRECT: Clear purpose and use cases
   /// <summary>
   /// Read, list and search mail items in the running Outlook session.
   /// </summary>
   ```

2. **Non-Enum Parameter Values Documented**:
   ```csharp
   // ❌ WRONG: Parameter values not explained
   /// <summary>List mail with a folder parameter</summary>
   
   // ✅ CORRECT: Non-enum parameter values explained
   /// <summary>
   /// List mail items in a folder.
   ///
   /// FOLDER ALIASES (non-enum parameter):
   /// - 'inbox': the default Inbox (DEFAULT)
   /// - 'drafts': the Drafts folder
   /// - 'sent': the Sent Items folder
   /// - 'deleted': the Deleted Items folder
   /// A full folder path is also accepted.
   /// </summary>
   ```

3. **Server-Specific Behavior Documented**:
   ```csharp
   // ❌ WRONG: Behavior changed but description outdated
   /// <summary>Default: shapeType='line'</summary>  // Wrong!
   
   // ✅ CORRECT: Description reflects actual default
   /// <summary>Default: shapeType='rectangle'</summary>
   ```

**What NOT to include:**
- ❌ **Enum action lists** - MCP SDK auto-generates these in schema (LLMs see them via dropdown)
- ❌ **Parameter types** - Schema provides this
- ❌ **Required/optional flags** - Schema provides this

**Why Critical:** LLMs use tool descriptions for server-specific guidance. Inaccurate descriptions cause:
- Wrong default parameter values
- Incorrect workflow assumptions
- Confused users when behavior doesn't match docs

**When to Update:**
- Changing default values or server behavior
- Adding/changing non-enum parameter values (shapeType, transitionType, etc.)
- Changing which tools to use for related operations
- Adding performance guidance (batch mode)

**See:** [mcp-server-guide.instructions.md](mcp-server-guide.instructions.md) for complete Tool Description checklist.

---

## Rule 19: Check PR Review Comments After Creating PR (CRITICAL)

**After creating a PR, ALWAYS check for automated review comments from Copilot and GitHub Advanced Security.**

```powershell
# Retrieve inline code review comments using GitHub CLI
# ⚠️ IMPORTANT: gh CLI requires authentication with a PERSONAL GitHub account.
# Enterprise Managed User (EMU) accounts cannot access public repos via gh CLI.
# Use: gh auth login --with-token (with a personal access token)
gh api repos/trsdn/mcp-server-outlook/pulls/PULL_NUMBER/comments --paginate

# Or use the mcp_github tool if available
mcp_github_github_pull_request_read(method="get_review_comments", owner="trsdn", repo="mcp-server-outlook", pullNumber=PULL_NUMBER)
```

**Common automated reviewers:**
- **Copilot** (code quality, performance, style)
- **github-advanced-security** (security scanning, code analysis)

**Common issues to fix:**
- Improper `/// <inheritdoc/>` on constructors/test methods that don't override
- `.AsSpan().ToString()` inefficiency - use `[..n]` range operator instead
- Nullable type access without null checks
- `foreach` → `.Select()` for functional style
- Nested if statements that can be combined
- Generic catch clauses - use specific exceptions or add justification
- Path.Combine security warnings - suppress with justification for test code

**Fix all automated review comments before requesting human review.**

**Why Critical:** Automated reviewers catch common code quality issues early. Fixing them promptly:
- Improves code quality and maintainability
- Reduces human reviewer workload
- Speeds up PR approval process
- Prevents accumulation of technical debt

**Process:**
1. Create PR
2. Immediately check for review comments (within 1-2 minutes)
3. Fix all automated issues in a single follow-up commit
4. Push fixes to PR branch
5. Request human review only after all automated issues resolved

**Example:** PR #139 had 17 automated review comments - all fixed in one commit before human review.

---

## Rule 22: COM Cleanup Must Use Finally Blocks (CRITICAL)

**ALWAYS use try-finally for COM object cleanup. NEVER swallow exceptions with empty catch blocks.**

```csharp
// WRONG: swallows the exception and sets a fallback value
try
{
    Outlook.MAPIFolder parent = folder.Parent;
    Outlook.NameSpace ns = parent.Session;
    name = ns.CurrentUser.Name;
    OutlookInteropRunner.ReleaseComObject(ref ns);      // Not reached if anything throws
    OutlookInteropRunner.ReleaseComObject(ref parent);
}
catch
{
    name = "(unknown)";  // Swallows the exception, including an OMG security denial
}

// CORRECT: finally guarantees cleanup, the exception propagates
Outlook.MAPIFolder? parent = null;
Outlook.NameSpace? ns = null;
try
{
    parent = folder.Parent;
    ns = parent.Session;
    name = ns.CurrentUser.Name;
}
finally
{
    OutlookInteropRunner.ReleaseComObject(ref ns);
    OutlookInteropRunner.ReleaseComObject(ref parent);
}
// Exception reaches the runner, COM objects are always released
```

**Why Critical:**
- Finally blocks execute **regardless** of exceptions
- COM objects leak if release is not reached before the exception
- Swallowing exceptions hides real problems. `"(unknown)"` is indistinguishable from a
  successful read of a mailbox literally named "(unknown)", and it silently hides Object Model
  Guard denials, which is exactly the case a caller most needs to know about
- Empty catch blocks are a code smell - remove them

**Pattern Requirements:**
1. Declare COM objects as nullable (`Outlook.MAPIFolder?`) before the try block
2. Initialize to `null`
3. Acquire COM objects inside the try block
4. Release in the finally block, children before parents
5. **NO catch blocks** unless a specific exception genuinely needs different handling
6. **NEVER** catch just to set a fallback value
7. Use `ReleaseSharedComObject` for the shared `Outlook.Application`, never `ReleaseComObject`

**See Also:**
- Rule 1b: Exception propagation pattern
- outlook-com-interop.instructions.md for complete patterns

---

## Rule 23: NEVER Dismiss IDE/Linter Warnings as False Positives (CRITICAL)

**When VS Code, linters, or other tooling shows errors or warnings, TRUST THEM. Do not dismiss them as "false positives" without verification.**

**Why Critical:** Dismissing valid warnings leads to broken code reaching production. The agent's job is to write correct code, not rationalize why incorrect code is acceptable.

```yaml
# ❌ WRONG: Agent dismissed YAML error as "false positive"
run: |
  $notes = @"
  ## Release Notes
  ```powershell
  code here
  ```
  "@
# VS Code showed: "Unexpected scalar at node end" - THIS WAS A REAL ERROR

# ✅ CORRECT: Agent should have tested or researched before dismissing
# PowerShell here-strings (@"..."@) don't work in GitHub Actions YAML
# Use string concatenation instead:
run: |
  $notes = "## Release Notes`n"
  $notes += '```powershell' + "`n"
  $notes += "code here`n"
  $notes += '```'
```

**Enforcement:**
- If IDE shows error/warning, assume it's CORRECT until proven otherwise
- To disprove a warning, you MUST either:
  1. Run the code and verify it works, OR
  2. Find authoritative documentation proving the warning is wrong
- "I think it will work" is NOT verification
- "The linter is confused" is NOT a valid dismissal

**Common False Positive Claims That Are Usually WRONG:**
- "YAML linter doesn't understand multi-line strings" - It usually does
- "PowerShell syntax is valid, YAML just can't parse it" - If YAML can't parse it, the workflow fails
- "This works locally" - GitHub Actions environment may differ

**What To Do Instead:**
1. See a warning → Pause and investigate
2. Research the specific syntax/pattern
3. Test if possible (run workflow, run code locally)
4. If warning is truly false positive, document WHY with evidence
5. If uncertain, use a simpler approach that doesn't trigger warnings

---

## Rule 24: Post-Change Sync Verification (CRITICAL)

**After adding or modifying any tool/action, ALWAYS verify ALL sync points are updated. This is MANDATORY before commit.**

**Sync Points Checklist:**

The MCP tool, its action enum, the action-to-string mapping and the CLI command are all
**generated** from the Core interface attributes by `src/OutlookMcp.Generators*`. There is no
`ToolActions.cs` and no `ActionExtensions.cs`; they were deleted when the generators landed
(#5/#11). You therefore cannot forget a switch case. What you can forget is everything the
generator does not produce.

When adding a NEW action to an existing tool:

| Sync Point | What to update |
|------------|----------------|
| 1. Interface | `I*Commands.cs` - add the method **with a `[ServiceAction]` attribute** |
| 2. Core | `*Commands.cs` - implement via `OutlookInteropRunner.Execute` |
| 3. Destructive flag | `[ServiceAction("...", Destructive = true)]` for anything that mutates or deletes |
| 4. XML summary | The `/// <summary>` becomes the MCP tool description. Document preconditions and Object Model Guard exposure. |
| 5. Build | `dotnet build -c Release` - the tool, enum and CLI command appear |
| 6. Coverage test | `dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"` |
| 7. Feature count | `FEATURES.md` - update the operation count |
| 8. README files | Every README carrying an operation count (main, MCP, CLI, mcpb, vscode) |
| 9. Skills docs | `skills/shared/*.md` - document the new action. This is the single source of truth for both skill references and the generated MCP prompts. |
| 10. CHANGELOG | Rule 27 |

**Quick Check Commands:**
```powershell
# Find every file carrying an operation count, so none is missed
Get-ChildItem -Recurse -Filter *.md | Select-String -Pattern '\d+ operations|\d+ ops'

# Confirm what the surface actually is - [ServiceCategory] in Core is the source of truth,
# NOT the generated files under obj/, which can be stale
Get-ChildItem src\OutlookMcp.Core\Commands -Recurse -Filter I*.cs |
    Select-String -Pattern 'ServiceCategory\("([^"]+)"\)'
```

**Why Critical:**
- Missing `[ServiceAction]` = the action never reaches either entry point
- Outdated READMEs = users misjudge what the server can do
- Skills docs out of sync = LLMs give wrong instructions
- A stale `obj/` will happily show you tools for a product this repository no longer automates

**When to Run This Checklist:**
- After adding ANY new action to ANY tool
- After adding ANY new tool
- After removing or deprecating actions
- Before EVERY commit that touches tool or action code

**Historical Example:**
A `duplicate` action was added to the interface, the implementation and the MCP handler, but
its CLI handler, the `FEATURES.md` count and the README counts were all missed. That was under
the old hand-written routing. The generators removed the first class of miss entirely; items
7 to 10 above are the ones that still bite, because no compiler checks them.

---

## Rule 25: Use PowerShell Syntax in Documentation (CRITICAL)

**OutlookMcp is Windows-only. ALL documentation code blocks MUST use PowerShell syntax, NOT bash.**

```markdown
# ❌ WRONG: bash syntax
```bash
dotnet build
outlookcli folder list-default
```

# ✅ CORRECT: PowerShell syntax
```powershell
dotnet build
outlookcli folder list-default
```
```

**Why Critical:**
- OutlookMcp requires Windows + a running classic Outlook for Windows
- bash syntax confuses Windows users
- PowerShell is the native Windows shell
- Syntax highlighting differs between bash/powershell

**Enforcement:**
- Use `powershell` or `pwsh` code fence, NEVER `bash` or `sh`
- Exception: Docker/Linux-specific docs (e.g., GitHub Actions runners on Linux)
- Review all new .md files for bash code blocks before commit

**Quick Check:**
```powershell
# Find all bash code blocks in markdown files
Select-String -Path "**/*.md" -Pattern '```bash' -Recurse
```

---

## Rule 27: Update CHANGELOG Before Merging PRs (CRITICAL)

**ALWAYS update CHANGELOG.md before merging any PR that adds features, fixes bugs, or makes breaking changes.**

**Why Critical:** 
- Users and LLMs rely on CHANGELOG to understand what changed
- Release workflow extracts version notes from CHANGELOG
- Missing entries make releases incomplete and confusing
- Git history is not a substitute for curated change documentation

**What Requires CHANGELOG Entry:**
- ✅ Bug fixes (with issue number)
- ✅ New features or actions
- ✅ Breaking changes
- ✅ Significant behavior changes
- ✅ New pre-commit checks or tooling
- ❌ Internal refactoring (no user-visible change)
- ❌ Documentation-only changes (unless significant)
- ❌ Test-only changes

**Format (Keep a Changelog):**
```markdown
## [Unreleased]

### Added
- **Feature Name** (#issue): Brief description

### Fixed
- **Bug Title** (#issue): Brief description
  - ROOT CAUSE: What caused it
  - FIX: How it was fixed

### Changed
- **Change Title**: Brief description
```

**Process:**
1. Before merging PR, check if changes need CHANGELOG entry
2. Add entry under `## [Unreleased]` section
3. Use appropriate category: Added, Fixed, Changed, Removed, Security
4. Include issue/PR number for traceability
5. Commit CHANGELOG update with the PR or as final commit before merge

**Enforcement:**
- Review CHANGELOG.md before every PR merge
- Agent must ask: "Does this change need a CHANGELOG entry?"
- Merging without CHANGELOG update for user-visible changes = bug

---

## Rule 28: COM API Parameter Naming (CRITICAL)

**When naming parameters on COM API wrapper methods, apply this nuanced principle:**

> **If the COM method's parameter name is clear and self-describing in our flat tool schema, use it.
> If the COM name is opaque or ambiguous without its parent context, keep a more descriptive name.**

**Why:** MCP tool parameters appear in a flat schema — they lose the context of the parent class/method. A name that works when you see `MailItem.Subject` may be opaque when the LLM just sees a `subject` parameter. Conversely, inventing a name like `subjectLine` when COM already calls it `Subject` adds unnecessary indirection.

**Decision Framework:**

| COM API | COM Param | Our Param | Rationale |
|---------|-----------|-----------|-----------|
| `Slides.Add(Index, Layout)` | `Index` | `slideIndex` | ✅ Keep descriptive — `index` alone is ambiguous |
| `MailItem.Subject` | `Subject` | `subject` | ✅ `subject` is self-describing in the tool schema |
| `MAPIFolder.Name` | `Name` | `folderName` | ✅ Keep descriptive — COM's `Name` property is too generic in a flat schema |
| `TextFrame.Text` | (property) | `text` | ✅ Clear in context |
| `SlideRange.Item(Index)` | `Index` | `slideIndex` | ✅ Keep descriptive — `index` alone is ambiguous |

**Implementation Pattern:**
```csharp
// ✅ COM name is clear → use it directly
OperationResult SetSubject(string entryId, string subject);

// ✅ COM name works in flat schema → use it
OperationResult SetBody(string entryId, string body);

// ✅ COM name too generic → keep descriptive
OperationResult Move(string entryId, string folderName, string? storeId = null);
```

**When Adding New Parameters:**
1. Check the COM API docs for the original parameter name
2. Ask: "Would an LLM understand `{com_param_name}` without seeing the method/class name?"
3. If YES → use COM name (e.g., `rotation`, `text`, `left`)
4. If NO → use descriptive name (e.g., `shapeName` not `Name`, `slideIndex` not `Index`)

---

## Rule 29: Test-Driven Development (TDD) (CRITICAL)

**Write the test FIRST. Watch it FAIL. Then implement. Then watch it PASS.**

This is non-negotiable for all new features and bug fixes. The workflow is:

1. **Write a failing test** that describes the expected behavior
2. **Run the test** — it MUST fail (red)
3. **Implement the minimum code** to make the test pass
4. **Run the test again** — it MUST pass (green)
5. **Refactor** if needed, ensuring tests still pass

```csharp
// Step 1: Write the test FIRST (before any implementation)
[Fact]
[Trait("Feature", "Progress")]
public void ProgressAdapter_Maps_Current_To_Progress()
{
    // This test describes what the code SHOULD do
    // It will FAIL because the class doesn't exist yet
    var reported = new List<ProgressNotificationValue>();
    IProgress<ProgressNotificationValue> inner = new Progress<ProgressNotificationValue>(v => reported.Add(v));
    var adapter = new McpProgressAdapter(inner);

    adapter.Report(new ProgressInfo { Current = 0.5f, Total = 1.0f, Message = "Half done" });

    Assert.Single(reported);
    Assert.Equal(0.5f, reported[0].Progress);
}

// Step 2: Run → RED (McpProgressAdapter doesn't exist)
// Step 3: Implement McpProgressAdapter
// Step 4: Run → GREEN
// Step 5: Refactor if needed
```

**Why Critical:**
- **Proves the test actually tests something** — a test that never fails never caught a bug
- **Defines behavior before implementation** — forces clear thinking about requirements
- **Prevents "tests that always pass"** — if your test passes before implementation, it's testing nothing
- **Catches regressions immediately** — the test existed before the code, so it's a true regression guard
- **Reduces debugging time** — small red→green cycles isolate problems instantly

**Common Violations:**
```
# ❌ WRONG: Write code first, then write tests that pass
1. Implement feature     ← No test to guide you
2. Write tests           ← Tests shaped to match implementation, not requirements
3. Tests pass            ← Of course they do — you wrote them to match existing code

# ✅ CORRECT: TDD cycle
1. Write failing test    ← Defines what "correct" means
2. Run test → RED        ← Proves the test catches the missing behavior
3. Implement feature     ← Guided by the test
4. Run test → GREEN      ← Proves the implementation is correct
```

**Enforcement:**
- When adding a new feature: write tests FIRST, commit them separately if needed
- When fixing a bug: write a test that reproduces the bug FIRST, then fix
- In PR reviews: verify that tests were written before or alongside implementation
- If a test passes on the first run without any implementation changes, investigate — it may be testing nothing

**Applies To:**
- All new Core Commands methods
- All new MCP Server tool actions
- All bug fixes (regression test first)
- All CLI command additions

**Exception:** Pure refactoring where existing tests already cover the behavior — no new tests needed, but existing tests must still pass.

---

## Rule 30: Integration Tests Over Unit Tests (CRITICAL)

**NEVER write unit tests that mock COM objects, fake contexts, or test adapter mappings in isolation — they prove NOTHING. Write integration tests that exercise real COM automation. A narrow exception for genuinely pure logic is defined below and enumerated in ADR-001.**

**Why Critical:** OutlookMcp is a COM interop project. The bugs that matter — STA threading deadlocks, RCW lifetime bugs, `OleMessageFilter` re-entrancy, Object Model Guard denials, type conversion failures — **only manifest when a real Outlook is running**. A unit test that verifies an adapter maps field A to field B catches zero real bugs. An integration test that reads a real folder and verifies the result catches all of them.

```csharp
// ❌ WRONG: Unit test that proves nothing
[Fact]
[Trait("Category", "Unit")]
public void Adapter_Maps_Field_A_To_Field_B()
{
    var adapter = new SomeAdapter(mockProgress);
    adapter.Report(new Info { Current = 0.5f });
    Assert.Equal(0.5f, captured.Progress);  // So what? This never fails in production.
}

// ✅ CORRECT: Integration test that catches real bugs
[Fact]
[Trait("Category", "Integration")]
[Trait("Feature", "Mail")]
public void List_Inbox_ReturnsItems_FromRunningOutlook()
{
    var result = _commands.List(folder: "inbox", maxResults: 5);

    Assert.True(result.Success);
    Assert.Null(result.ErrorMessage);
    Assert.NotEmpty(result.Items);  // Real Outlook, real store, real marshalling
}
```

**What Counts as Integration:**
- ✅ Reaches a real running Outlook via COM
- ✅ Exercises `OutlookInteropRunner.Execute` on the real dispatcher STA thread
- ✅ Verifies real data flows through the full pipeline
- ✅ Catches COM threading, RCW lifetime, marshalling and Object Model Guard bugs

**What Does NOT Count:**
- ❌ Mocking `Outlook.Application`, `Outlook.NameSpace`, or any COM interface
- ❌ Testing adapter/mapper classes in isolation
- ❌ Verifying AsyncLocal behaviour without a COM context
- ❌ Any test of COM behaviour that passes without OUTLOOK.EXE running

**Enforcement:**
- Code review MUST reject unit tests for COM-dependent features
- Any test of COM-dependent behaviour MUST have `[Trait("Category", "Integration")]`
- If a test doesn't require Outlook, question whether it tests anything meaningful

**The narrow exception (amended 2026-09-02, #37).** A unit test is permitted only if it meets **all four** conditions:

1. It touches **no COM object at all** — not a real one, not a `null!` stand-in, not a mock
2. Its subject is genuinely pure: a mapping, parse, classification, invariant, or a guard clause that returns before any COM call
3. It would fail if the logic under test were wrong (not merely if .NET itself were broken)
4. It carries `[Trait("Category", "Unit")]` and lives under `tests/**/Unit/`

If a test needs a COM object to mean anything and you are substituting something for that object to make it run, it is prohibited. Write an integration test, or write none.

**A permitted unit test never counts as coverage for a COM operation.** It proves only its own narrow claim.

`docs/ADR-001-NO-UNIT-TESTS.md` holds the normative list of files currently permitted under this exception. Adding to that list requires updating the ADR in the same PR.

**Historical Lesson:** 10 unit tests were written for the MCP progress feature (McpProgressAdapter mapping, ProgressContext AsyncLocal). All 10 passed. Zero of them would have caught the real bugs: STA thread affinity issues, COM callback re-entrancy during a live operation, or notifications not flowing through the generated code pipeline. The unit tests tested the unit tests.

**Second lesson (#37):** the rule as originally written said "never", while 16 unit test files sat in the repository. Three of them mocked COM outright — one was named `Release_WithComObject_DoesNotThrow` above a comment conceding no COM object was used. An absolute rule that the codebase visibly contradicts gets ignored wholesale rather than obeyed selectively, which is why the exception is now stated explicitly and its members enumerated.

