---
applyTo: "src/**/*.cs"
---

# Architecture Patterns

> **Core patterns for OutlookMcp development**

## .NET Class Design (MANDATORY)

**Official Docs:** [Framework Design Guidelines](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/), [Partial Classes](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods)

### Key Rules

1. **One Public Class Per File** - Standard .NET practice (System.Text.Json, ASP.NET Core, EF Core)
2. **File Name = Class Name** - `RangeCommands.cs` contains `RangeCommands`
3. **Partial Classes for Large Implementations** - Split 15+ method classes by feature domain
4. **Descriptive Names** - No over-optimization (`RangeCommands` ✅, `Commands` ❌)
5. **Folder = Organization, Not Identity** - `Commands/Range/RangeCommands.cs`

### Partial Class Pattern

**When:** Class has 15+ methods, multiple feature domains, team collaboration

**Structure:**
```
Commands/Range/
    IRangeCommands.cs           # Interface
    RangeCommands.cs            # Partial (constructor, DI)
    RangeCommands.Values.cs     # Partial (Get/Set values)
    RangeCommands.Formulas.cs   # Partial (formulas)
    RangeHelpers.cs             # Separate helper class
```

**Benefits:** Git-friendly, team-friendly, ~100-200 lines per file, mirrors .NET Framework patterns

---

## TWO EQUAL ENTRY POINTS (CRITICAL)

**OutlookMcp has TWO first-class entry points: MCP Server AND CLI.** Both must have:
- **Feature parity**: Every action in MCP must exist in CLI and vice versa
- **Parameter parity**: Same parameters, same defaults, same validation
- **Behavior parity**: Same Core command, same result format

When adding or changing ANY feature, ALWAYS update BOTH entry points. See Rule 24 (Post-Change Sync).

```
MCP Server (MCP tools, JSON-RPC) --> In-process OutlookMcpService --> Core Commands --> OutlookDispatcher --> Outlook COM
CLI (command-line args, console)  --> CLI Daemon (named pipe) -----> Core Commands --> OutlookDispatcher --> Outlook COM
```

Both entry points funnel into the same single process-wide `OutlookDispatcher` STA thread.

---

## Command Pattern

### Structure
```
Commands/Mail/
  IMailCommands.cs   # Interface, annotated with [ServiceCategory] / [ServiceAction]
  MailCommands.cs    # Implementation
```

### Routing

You do not hand-write routing. `src/OutlookMcp.Generators*` reads the `[ServiceCategory]` and
`[ServiceAction]` attributes on the Core interface and generates both the CLI command and the
MCP tool action. Adding an action means adding an attributed interface method and implementing
it; the surface follows automatically, which is what keeps CLI/MCP parity structural rather
than manual.

---

## Resource Management Pattern

**See outlook-com-interop.instructions.md** for the `OutlookInteropRunner.Execute` pattern and
COM object lifecycle management.

---

## Exception Propagation Pattern (CRITICAL)

**Core Commands: the `onException` delegate is the only place an exception becomes a failed
result.** Do not add a `catch` inside the action lambda that returns an error result.

```csharp
// WRONG: a catch inside the action that returns a failure result.
// It shadows the runner's Object Model Guard classification and loses the real cause.
return OutlookInteropRunner.Execute(
    "OutlookMailRead",
    (application, session) =>
    {
        try { /* ... */ }
        catch (Exception ex)
        {
            return new ActiveMailResult { Success = false, ErrorMessage = ex.Message };
        }
    },
    ex => new ActiveMailResult { Success = false, ErrorMessage = ex.Message });

// CORRECT: let it propagate to onException
return OutlookInteropRunner.Execute(
    "OutlookMailRead",
    (application, session) =>
    {
        Outlook.MailItem? mail = null;
        try
        {
            mail = /* ... */;
            return new ActiveMailResult { Success = true /* ... */ };
        }
        finally
        {
            // Finally blocks are required for COM cleanup and are unaffected by this rule
            OutlookInteropRunner.ReleaseComObject(ref mail);
        }
    },
    ex => new ActiveMailResult
    {
        Success = false,
        ErrorMessage = $"Failed to read the active mail item: {ex.Message}"
    });
```

**Why this pattern:**
- `OutlookInteropRunner.Execute` already classifies COM failures, Object Model Guard denials,
  "Outlook not running" and elevation mismatches into actionable messages
- A local catch replaces those messages with a generic one and originates from the wrong layer
- `finally` is the correct place for resource cleanup, `catch` is not the place for error
  suppression

**See:** critical-rules.instructions.md Rule 1 for Success flag requirements

---

## MCP Server Tools

**In-process architecture**: the MCP Server hosts `OutlookMcpService` fully in-process with
direct method calls (no pipe). `ServiceBridge` holds the service reference and calls
`ProcessAsync()` directly.

**Five generated tools**, one per `[ServiceCategory]`:

1. `application` - Outlook application state and diagnostics
2. `attachment` - attachment listing and extraction
3. `calendar` - calendar and appointment operations
4. `folder` - folder navigation and listing
5. `mail` - mail read, list, search, compose

Each tool takes an `action` parameter whose values come from the generated action enum. Both
the tool and the enum are produced from the Core interface attributes, so they cannot drift
out of sync with Core. `CoreCommandsCoverageTests` asserts this by reflection in CI.

---

## Security-First Patterns

Outlook's Object Model Guard blocks out-of-process access to protected members such as
`SenderEmailAddress`, `Recipients` and `MailItem.Send()`. This server never attempts to bypass
it - no registry trust edits, no security-manager tricks. When an operation touches a
protected member, document that in its XML summary so the generated tool description warns the
caller. See `docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md`.

Destructive actions are declared explicitly: `[ServiceAction("...", Destructive = true)]`
propagates to the MCP tool annotation so clients can prompt before running them.

---

## Performance Patterns

**Prefer server-side filtering.** `Items.Restrict`, `Items.Find` and `Table` push the predicate
into Outlook's store. A client-side loop over `Items` marshals every item across the COM
boundary and does not scale past a few hundred (see #27).

**Minimize COM round-trips.** Every property read on a COM object is a cross-apartment call.
Read what you need once into a local, then release.

---

## Key Principles

1. **`OutlookInteropRunner.Execute` for everything** - see outlook-com-interop.instructions.md
2. **Release intermediate objects** in `finally`, children before parents
3. **Never final-release the shared `Outlook.Application`** - see #19
4. **Generated surface** - annotate the Core interface, do not hand-write tools or commands
5. **Security defaults** - never work around the Object Model Guard
6. **Server-side filtering** - minimize COM round-trips
