# Contributing to OutlookMcp

Thank you for your interest in contributing to OutlookMcp. The project is now focused on automating the classic Outlook for Windows desktop app through COM for MCP clients and the `outlookcli` CLI.

## Project Vision

OutlookMcp aims to provide a small, reliable Outlook automation surface for AI assistants and scripts:

- Mail operations.
- Calendar appointment operations.
- Folder discovery and item listing.
- Attachment inspection and mutation.
- Outlook application status checks.

The active product surface is 5 tools with 30 operations. Deleted presentation command domains are not part of the current product.

## Getting Started

### Development Environment

1. Prerequisites:
   - Windows OS.
   - .NET 10 SDK.
   - Visual Studio 2022 or VS Code.
   - Classic Outlook for Windows desktop app for manual Outlook testing.

2. Setup:

```powershell
git clone https://github.com/trsdn/mcp-server-outlook.git
Set-Location mcp-server-outlook
dotnet restore
dotnet build --nologo -v q
```

3. Test your setup:

```powershell
.\src\OutlookMcp.CLI\bin\Debug\net10.0-windows\outlookcli.exe application get-status
```

Classic Outlook must be installed, running, signed in, and at the same elevation level as the CLI process.

## CRITICAL: Pull Request Workflow Required

All changes must be made through pull requests. Direct commits to `main` are prohibited.

### Quick PR Process

1. Create feature branch: `git checkout -b feature/your-feature`.
2. Make changes: code, tests, and documentation.
3. Push branch: `git push origin feature/your-feature`.
4. Create PR using GitHub's PR template.
5. Address review.
6. Merge after approval and required checks pass.

See [DEVELOPMENT.md](DEVELOPMENT.md) for complete instructions.

## Development Guidelines

### Code Style

- Use modern C# patterns already present in the codebase.
- Nullable reference types are enabled. Handle nulls explicitly.
- Build with zero warnings.
- Use XML documentation for public APIs and tool-facing behavior.
- Keep MCP and CLI behavior in sync.
- Do not include private mailbox data, secrets, or customer identifiers in code, tests, commits, issues, or PRs.

### Architecture Patterns

#### Active Outlook Command Pattern

Active commands live under `src\OutlookMcp.Core\Commands\` and are exposed through interfaces marked with `[ServiceCategory]` and methods marked with `[ServiceAction]`.

Current categories:

- `mail`
- `calendar`
- `folder`
- `attachment`
- `application`

Example pattern:

```csharp
[ServiceCategory("application")]
[NoSession]
public interface IApplicationCommands
{
    [ServiceAction("get-status")]
    OutlookApplicationStatusResult GetStatus(bool includeActiveContext = true);
}
```

The generators use these interfaces to keep MCP tools and CLI commands aligned.

#### Outlook COM Execution

Outlook automation goes through `OutlookDispatcher` on a dedicated STA thread. See [ADR-002](ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md).

Key rules:

1. Use classic Outlook through `Outlook.Application`.
2. Do not add support claims for new Outlook; it has no COM object model.
3. Treat Outlook as a shared desktop application and real mailbox.
4. Use entry IDs from list, search, or read-active results for follow-up operations.
5. Re-resolve entry IDs after moves between stores.
6. Respect Outlook Object Model Guard prompts. Do not bypass them.

#### No Session or Batch API

Outlook has no document to open, save, or close. Every COM call is marshalled onto the single STA thread owned by `OutlookDispatcher`. Do not introduce a session or batch abstraction into an Outlook command.

### Testing

Before submitting:

1. Run `dotnet build --nologo -v q`.
2. Run targeted tests for the changed area.
3. If changing generated service actions, run `CoreCommandsCoverageTests`.
4. If changing Outlook behavior, manually test on Windows with classic Outlook running and document the result.
5. If you cannot run Outlook behavior tests, state that gap plainly.

Useful commands:

```powershell
dotnet build --nologo -v q
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
dotnet test tests\OutlookMcp.CLI.Tests
dotnet test tests\OutlookMcp.Core.Tests --filter "Feature=OutlookSeed"
```

Hosted CI does not currently verify Outlook behavior because a self-hosted Windows runner with classic Outlook is not available.

## Adding New Commands

### 1. Update or Add a Core Interface

Add the operation to the appropriate `[ServiceCategory]` interface, or add a new category only when the product surface requires a new tool.

```csharp
[ServiceAction("new-action")]
SomeResult NewAction(string requiredValue, bool optionalFlag = false);
```

### 2. Implement the Command

Implement the behavior in the matching Core command class. Use existing Outlook interop helpers and return result objects with accurate success and error state.

### 3. Rebuild Generated Code

```powershell
dotnet build --nologo -v q
```

### 4. Verify MCP and CLI Parity

```powershell
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
outlookcli mail --help
```

Use the relevant command instead of `mail` for other categories.

### 5. Update Documentation and Skills

Update user-facing docs, skills, and examples if the public command surface changes.

**`skills/outlook-cli/SKILL.md` and `skills/outlook-mcp/SKILL.md` are generated.** A Release build
regenerates both from `skills/templates/*.sbn` and `skills/shared/*.md`, so editing them directly
is silently undone the next time anyone runs `dotnet build -c Release`. Edit the template or the
shared guidance instead:

| To change | Edit |
| --------- | ---- |
| Skill structure, headings, transitional notes | `skills\templates\SKILL.cli.sbn`, `skills\templates\SKILL.mcp.sbn` |
| Behavioural guidance and workflows | `skills\shared\*.md` (also becomes the MCP prompts) |
| Generated `SKILL.md` files | nothing - they are build output |

After editing, run `dotnet build -c Release` and confirm `git status` is clean apart from the
regenerated `SKILL.md` files you intended to change.

## Pull Request Process

### Before Submitting

- [ ] Code builds with zero warnings.
- [ ] Targeted tests pass.
- [ ] MCP and CLI entry points remain in sync.
- [ ] Outlook behavior was manually tested when applicable, or the gap is stated.
- [ ] Documentation was updated as needed.
- [ ] No private mailbox data, secrets, or customer identifiers are included.

### PR Description Template

```markdown
## Summary
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] dotnet build --nologo -v q
- [ ] Targeted tests listed here
- [ ] Manual Outlook test listed here, or reason it was not run

## Checklist
- [ ] MCP and CLI parity preserved
- [ ] Documentation updated
- [ ] No sensitive data included
```

## UI Guidelines

### Spectre.Console Usage

```csharp
AnsiConsole.MarkupLine($"[green]Success:[/] Operation succeeded");
AnsiConsole.MarkupLine($"[red]Error:[/] {message.EscapeMarkup()}");
AnsiConsole.MarkupLine($"[yellow]Note:[/] {message.EscapeMarkup()}");
AnsiConsole.MarkupLine($"[dim]{message.EscapeMarkup()}[/]");
```

### Output Consistency

- Prefer JSON for automation-facing output.
- Escape markup when writing user-controlled text through Spectre.Console.
- Provide actionable errors for missing Outlook, new Outlook only, elevation mismatch, and Object Model Guard prompts.

## Bug Reports

Open a bug through the **Bug report** form. It requires the fields that make a report actionable:
`application.get-status` output, reproduction steps, expected and actual behaviour, and your
environment. Blank issues are disabled, so pick the form that matches what you are reporting - use
**MCP Server issue** when an AI assistant is driving the server, since that form also asks for your
client configuration and the exact tool and action.

Do not attach or paste private mailbox data. Redact subjects, addresses, bodies, attachment names,
and entry IDs.

Never open a public issue for a security vulnerability. Follow [SECURITY.md](../SECURITY.md).

## Feature Requests

Use the **Feature request** form. It asks for the problem before the solution, the Outlook domain,
and a safety class, because Outlook automation acts on a real mailbox and is frequently
irreversible. If a proposal is not read-only, say what confirmation or idempotency it needs.

Remember that the MCP and CLI surfaces are generated from the same interfaces, so a feature reaches
both. Proposals for only one surface will be sent back.

## Dependency Updates

Dependabot is configured in [`.github/dependabot.yml`](../.github/dependabot.yml) for the three
ecosystems that actually exist here: NuGet at the repo root, GitHub Actions, and npm under
`vscode-extension/`. If you add a manifest for a new ecosystem, add it to that file in the same PR -
otherwise Dependabot will raise alerts for it that no update PR ever arrives to fix.

@trsdn triages and merges Dependabot PRs. Security updates are merged as soon as CI is green;
grouped minor and patch updates are reviewed weekly. A Dependabot PR must pass the same checks as
any other PR.

Do not hand-edit `package-lock.json` to bump a transitive dependency, and be careful running
`npm install` behind a corporate registry mirror: it rewrites `resolved` URLs to the mirror and can
downgrade `integrity` from sha512 to sha1, which breaks installs for everyone else. Let Dependabot
produce the lockfile change.

## Learning Resources

- [Outlook Object Model Reference](https://learn.microsoft.com/office/vba/api/overview/outlook)
- [.NET COM Interop Guide](https://learn.microsoft.com/dotnet/standard/native-interop/cominterop)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Spectre.Console Documentation](https://spectreconsole.net/)

## For Maintainers

- [NuGet Publishing Guide](NUGET-GUIDE.md) - publishing packages with OIDC trusted publishing.

## Issue Labels

- `bug` - Something is not working.
- `enhancement` - New feature or improvement.
- `documentation` - Documentation improvements.
- `good first issue` - Good for newcomers.
- `help wanted` - Extra attention needed.
- `outlook-com` - Classic Outlook COM automation issues.
- `coding-agent` - Coding agent related.

---

Thank you for contributing to OutlookMcp.
