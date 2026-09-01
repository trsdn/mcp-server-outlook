---
applyTo: "src/OutlookMcp.Core/Commands/**/*.cs,src/OutlookMcp.McpServer/**/*.cs"
---

# Core Commands Coverage - Mandatory Workflow

> **⚠️ CRITICAL**: When adding Core Commands methods, you MUST expose them in MCP Server

## Quick Reference

| Task | Command | Time |
|------|---------|------|
| Check coverage before commit | `dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"` | 30s |
| Add new Core method | Follow 8-step workflow below | 5-10 min |
| Fix pre-commit hook failure | Add missing enum values + mappings | 2-3 min |
| Verify build | `dotnet build -c Release` | 1-2 min |

---

## Mandatory Workflow: Adding New Core Method

**ALWAYS follow these 8 steps in order:**

```markdown
1. ✅ Add method to Core Commands interface
   File: src/OutlookMcp.Core/Commands/[Feature]/I[Feature]Commands.cs
   Example: Task<OperationResult> NewMethodAsync(IPptBatch batch);

2. ✅ Implement in Core Commands class  
   File: src/OutlookMcp.Core/Commands/[Feature]/[Feature]Commands.cs

3. ✅ Add enum value to ToolActions.cs
   File: src/OutlookMcp.McpServer/Models/ToolActions.cs
   Example: SlideAction.NewMethod
   ⚠️ Build will show CS8524 error until steps 4-6 complete

4. ✅ Add ToActionString mapping
   File: src/OutlookMcp.McpServer/Models/ActionExtensions.cs
   Example: SlideAction.NewMethod => "new-method",
   ⚠️ CS8524 error persists

5. ✅ Add switch case in MCP Tool
   File: src/OutlookMcp.McpServer/Tools/Ppt[Feature]Tool.cs
   Example: SlideAction.NewMethod => await NewMethodAsync(...),
   ⚠️ CS8524 error persists

6. ✅ Implement MCP method
   File: src/OutlookMcp.McpServer/Tools/Ppt[Feature]Tool.cs
   Example: private static async Task<string> NewMethodAsync(...)
   ✅ CS8524 errors resolved

7. ✅ Build and verify
   Command: dotnet build -c Release
   Expected: 0 warnings, 0 errors

8. ✅ Update documentation
   Files: skill references (`skills/shared/`), tool descriptions, README (if needed)
```

**Why This Order**: Compiler (CS8524) enforces steps 3-6, preventing you from shipping unexposed Core methods.

---

## Compiler Enforcement (CS8524)

**The compiler FORCES you to expose Core methods** through enum-based switches:

```csharp
// Step 3: Add enum value (compiler checks this)
public enum SlideAction
{
    List,
    Get,
    NewMethod  // ⚠️ Forget this → CS8524 error in ActionExtensions.cs
}

// Step 4: Add ToActionString mapping (compiler checks this)
public static string ToActionString(this SlideAction action) => action switch
{
    SlideAction.List => "list",
    SlideAction.Get => "get",
    SlideAction.NewMethod => "new-method",  // ⚠️ Forget this → CS8524 error
};

// Step 5: Add switch case in Tool (compiler checks this)
return action switch
{
    SlideAction.List => await ListAsync(...),
    SlideAction.Get => await GetAsync(...),
    SlideAction.NewMethod => await NewMethodAsync(...),  // ⚠️ Forget this → CS8524 error
};
```

**Result**: **Impossible to compile** until all 3 enum mappings are added!

---

## Pre-Commit Hook (Automatic Check)

**Before every commit**, the pre-commit hook runs `CoreCommandsCoverageTests` (a reflection-based
xUnit test, not a PowerShell script -- see #25) to verify Core methods match generated enum
action values.

**Setup** (one-time):
```powershell
.\scripts\pre-commit.ps1
```

**On failure, you see an xUnit assertion failure like**:
```
IMailCommands_AllMethodsHaveEnumValues [FAIL]
  IMailCommands has 18 [ServiceAction] methods but MailAction has only 16 enum values.
```

**Fix**: Follow 8-step workflow above.

**Emergency bypass** (use only for non-Core changes):
```bash
git commit --no-verify -m "Message"
```

⚠️ **Never use `--no-verify`** for Core Commands changes - fix the gaps instead!

---

## Manual Coverage Check

**Run anytime** to verify coverage:

```powershell
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

**Expected output when coverage is complete**: all tests pass (each `I{X}Commands_AllMethodsHaveEnumValues`
fact asserts `enumValueCount >= coreMethodCount` for one Outlook Core interface).

**When gaps detected**: the specific interface's fact fails with a message naming the interface,
its Core method count, and the enum's (smaller) value count.

**Fix**: Follow 8-step workflow.

---

## Troubleshooting

### CS8524 Error: "Switch expression does not handle all possible values"

**Cause**: Added enum value but forgot to add it to switch expression.

**Fix**: Add the missing case to the switch expression in the file mentioned in error.

### Pre-Commit Hook Fails with "Coverage gaps detected"

**Cause**: Core interface has more methods than corresponding enum has values.

**Fix**: Follow 8-step workflow (steps 3-6).

### Build Succeeds but Pre-Commit Hook Still Fails

**Cause**: Added Core method but forgot to add enum value.

**Fix**: Add to ToolActions.cs, then mappings in ActionExtensions.cs, then Tool switch case.

---

## Key Takeaways

✅ **Compiler enforces coverage** - CS8524 prevents incomplete implementations  
✅ **Pre-commit hook verifies** - Catches gaps before commit  
✅ **8-step workflow is mandatory** - No shortcuts  
✅ **100% coverage is required** - No exceptions

