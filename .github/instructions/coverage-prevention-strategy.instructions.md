---
applyTo: "src/OutlookMcp.Core/Commands/**/*.cs,src/OutlookMcp.McpServer/**/*.cs"
---

# Core Commands Coverage - Mandatory Workflow

> **When you add a Core Commands method, it must reach both the MCP Server and the CLI.**

## Quick Reference

| Task | Command | Time |
|------|---------|------|
| Check coverage before commit | `dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"` | 30s |
| Add a new Core method | Follow the workflow below | 5-10 min |
| Verify build | `dotnet build -c Release` | 1-2 min |

---

## The surface is generated, not hand-written

This is the single most important thing to know. `src/OutlookMcp.Generators*` reads the
`[ServiceCategory]` and `[ServiceAction]` attributes on the Core interfaces and generates:

- the MCP tool and its `action` enum
- the action-to-string mapping
- the CLI command and its argument binding

`src/OutlookMcp.McpServer/Tools/` contains exactly one hand-written file, `OutlookToolsBase.cs`.
There is no `ToolActions.cs` and no `ActionExtensions.cs` - they were deleted when the
generators landed (#5/#11). If a guide tells you to edit them, that guide is stale.

You therefore cannot ship an unexposed Core method by forgetting a switch case, because there
is no switch case for you to forget. What you *can* forget is the attribute.

---

## Workflow: adding a new Core method

```markdown
1. Add the method to the Core Commands interface, WITH a [ServiceAction] attribute
   File: src/OutlookMcp.Core/Commands/[Domain]/I[Domain]Commands.cs
   Example: [ServiceAction("archive", Destructive = true)]
            OperationResult Archive(string entryId, string? storeId = null);

   The [ServiceCategory("...")] attribute is on the interface itself and already exists.
   Set Destructive = true for anything that mutates or deletes user data - it propagates
   to the MCP tool annotation so clients can prompt.

2. Implement it in the Core Commands class
   File: src/OutlookMcp.Core/Commands/[Domain]/[Domain]Commands.cs
   Route through OutlookInteropRunner.Execute. See outlook-com-interop.instructions.md.

3. Build
   Command: dotnet build -c Release
   Expected: 0 warnings, 0 errors. The tool, the enum, and the CLI command all appear.

4. Verify coverage
   Command: dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"

5. Update documentation
   Files: skill references (skills/shared/), the method's XML summary, README if the
   operation count changed. See Rule 24 (Post-Change Sync).
```

Parameter naming matters here: the generator calls `StringHelper.ToSnakeCase()` on your C#
parameter name to produce the MCP parameter name. Never put an underscore in the C# name.
Choose a camelCase name that produces the snake_case you want (`entryId` -> `entry_id`), or
use `[FromString("desiredName")]` if it cannot.

---

## Pre-commit hook (automatic check)

The pre-commit hook runs `CoreCommandsCoverageTests`, a reflection-based xUnit test rather
than a PowerShell script (see #25), to verify that every `[ServiceAction]` method on a Core
interface has a corresponding generated enum value.

**On failure you see an xUnit assertion like:**
```
IMailCommands_AllMethodsHaveEnumValues [FAIL]
  IMailCommands has 18 [ServiceAction] methods but MailAction has only 16 enum values.
```

In practice this almost always means one of two things: the method is missing its
`[ServiceAction]` attribute, or the generator output is stale. Source generators do not clean
stale outputs, so a `dotnet clean` followed by a rebuild resolves the second case.

**Emergency bypass** (only for changes that touch no Core interface):
```powershell
git commit --no-verify -m "Message"
```

Never use `--no-verify` for a Core Commands change. Fix the gap.

---

## Do not read `obj/` to learn what the surface is

Generated files under `src/OutlookMcp.McpServer/obj/` can contain output from a previous
build, including tools for domains that no longer exist. The authoritative answer to "what
tools does this server expose" is the set of `[ServiceCategory]` attributes in
`src/OutlookMcp.Core/Commands/`. Today that is five: `application`, `attachment`, `calendar`,
`folder`, `mail`.

---

## Troubleshooting

### The coverage test fails but the build succeeded
Stale generator output. `dotnet clean`, then rebuild.

### A new method does not appear in the CLI or MCP tool
It is missing its `[ServiceAction]` attribute, or it is on the implementation class rather
than the interface. The generator reads the interface.

### MSB3021 / MSB3027 file-lock errors during build
Not a code error. Something is holding the output assemblies - most often the `outlookcli`
daemon, which auto-starts on any CLI invocation. The locking PID is named in the MSB3027
message; stop that process and rebuild.

---

## Key Takeaways

- **The surface is generated** - annotate the interface, do not hand-write tools or commands
- **`CoreCommandsCoverageTests` verifies it** - in the pre-commit hook and in CI
- **`obj/` is not the source of truth** - `[ServiceCategory]` in Core is
- **Both entry points, always** - MCP and CLI parity is structural, keep it that way
