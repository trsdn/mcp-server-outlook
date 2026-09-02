# Development Workflow

## IMPORTANT: All Changes Must Use Pull Requests

Direct commits to `main` are not allowed. All changes must go through the pull request process to ensure:

- Code review and quality control.
- Proper version management.
- CI validation for the parts CI can actually run.
- Documentation updates.

## Standard Development Workflow

### 1. Create Feature Branch

```powershell
git checkout -b feature/your-feature-name
git checkout -b fix/issue-description
git checkout -b docs/update-description
```

### 2. Make Your Changes

```powershell
git add .
git commit -m "Describe the change clearly"
```

Do not include private mailbox data, customer names, secrets, or local file names from real user data in commits, issues, or PR descriptions.

### 3. Push Feature Branch

```powershell
git push origin feature/your-feature-name
```

### 4. Create Pull Request

1. Go to [GitHub Repository](https://github.com/trsdn/mcp-server-outlook).
2. Click "New Pull Request".
3. Select your feature branch.
4. Fill out the PR template with:
   - Clear title.
   - What changed and why.
   - Testing performed.
   - Known gaps, especially if Outlook behavior could not be tested.
   - Documentation updates.

### 5. PR Review Process

- Automated checks run.
- Maintainers review.
- Address feedback.
- Merge after approval and passing required checks.

### 6. After Merge

```powershell
git checkout main
git pull origin main
git branch -d feature/your-feature-name
git push origin --delete feature/your-feature-name
```

## Release Process

### Creating a New Release

Only maintainers can create releases.

1. Ensure all changes are merged to `main` through PRs.
2. Ensure `CHANGELOG.md` has the appropriate release notes.
3. Create and push a semantic version tag:

```powershell
git tag v1.1.0
git push origin v1.1.0
```

4. Let the release workflow build packages and publish artifacts.

### Version Numbering

We follow [Semantic Versioning](https://semver.org/):

- Major: breaking changes.
- Minor: backward-compatible features.
- Patch: backward-compatible fixes.

## Branch Protection Rules

The `main` branch should be protected with:

- Required pull request reviews.
- Required status checks.
- Up-to-date branches before merge.
- No direct pushes.

## Testing Requirements and Organization

### Current Reality

OutlookMcp's active runtime surface automates classic Outlook for Windows through COM. Hosted CI does not verify Outlook behavior today because that requires a self-hosted Windows runner with classic Outlook installed, running, signed in, and licensed.

CI-safe tests can still verify generated metadata, protocol plumbing, serialization, CLI daemon behavior, and pure .NET validation.

### Test Layout

```text
tests\
|-- OutlookMcp.Core.Tests\          # Core models, helpers, and manual Outlook smoke tests
|-- OutlookMcp.McpServer.Tests\     # MCP protocol and generated action coverage
|-- OutlookMcp.CLI.Tests\           # CLI routing, daemon, diagnostics, and batch command tests
|-- OutlookMcp.ComInterop.Tests\    # Shared COM infrastructure and dormant legacy session layer
|-- OutlookMcp.Diagnostics.Tests\   # Manual diagnostics when present
`-- OutlookMcp.SkillGeneration.Tests\ # Skill markdown quality checks
```

### Development Workflow Commands

During development:

```powershell
dotnet build --nologo -v q
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

For CLI work:

```powershell
dotnet test tests\OutlookMcp.CLI.Tests
```

For manual Outlook smoke coverage:

```powershell
dotnet test tests\OutlookMcp.Core.Tests --filter "Feature=OutlookSeed"
```

For dormant legacy session infrastructure:

```powershell
dotnet test tests\OutlookMcp.ComInterop.Tests --filter "Feature=PptBatch|Feature=PptSession|Feature=SessionManager"
```

### Adding New Tests

When changing the active Outlook command surface:

1. Add tests for generated action coverage or pure validation where possible.
2. Add Outlook smoke or integration coverage when behavior depends on real Outlook COM.
3. Mark Outlook-dependent tests with `RequiresOutlook=true`.
4. Make clear in the PR whether the Outlook-dependent test was run manually.
5. Never use fake Outlook behavior to claim COM behavior is verified.

Example trait set for real Outlook behavior:

```csharp
[Trait("Category", "Integration")]
[Trait("Feature", "OutlookSeed")]
[Trait("RequiresOutlook", "true")]
```

### PR Testing Requirements

Before creating a PR, run the smallest relevant validation:

```powershell
dotnet build --nologo -v q
```

Then add targeted tests for the changed area. If you changed Core service interfaces, run:

```powershell
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

If you could not run Outlook behavior tests because the required desktop setup is missing, state that plainly in the PR.

## Source Agent Client

The repository may still contain source-side agent components from the pre-Outlook migration. Treat those workflows as historical until they are verified against the current Outlook surface.

Before changing `src\OutlookMcp.Agent\**`, inspect its README and code for current behavior. Do not assume it is part of the released Outlook product.

## CLI Command Code Generation

### Architecture Overview

The active Outlook commands are defined by Core interfaces marked with `[ServiceCategory]` and `[ServiceAction]`.

Current service categories:

- `mail`
- `calendar`
- `folder`
- `attachment`
- `application`

Source generators produce:

- Action enums and action-string mappings.
- CLI settings and command routing.
- MCP-visible service metadata.
- CLI command registration for generated service commands.

### Adding or Changing a Service Action

1. Update the relevant Core interface and implementation.
2. Rebuild so generated code updates.
3. Verify the MCP server surface.
4. Verify the `outlookcli` command and help output.
5. Run `CoreCommandsCoverageTests`.
6. Update user documentation and skills if the public surface changed.

```powershell
dotnet build --nologo -v q
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
outlookcli mail --help
```

### Parity Rule

MCP server tools and `outlookcli` are equal entry points. Every active Outlook operation must be available through both, with the same behavior and parameters.

## MCP Server Configuration Management

Generated tool metadata is the source of truth for the active Outlook surface. Do not hand-document deleted presentation tools.

When MCP behavior changes, verify:

```powershell
dotnet build src\OutlookMcp.McpServer\OutlookMcp.McpServer.csproj --nologo -v q
dotnet test tests\OutlookMcp.McpServer.Tests --filter "FullyQualifiedName~CoreCommandsCoverageTests"
```

## PR Template Checklist

When creating a PR, verify:

- [ ] Code builds with zero warnings.
- [ ] Targeted tests were run.
- [ ] Outlook behavior was manually tested when the change depends on Outlook COM, or the gap is stated.
- [ ] MCP and CLI entry points remain in sync.
- [ ] Documentation was updated if public behavior changed.
- [ ] Breaking changes are documented.
- [ ] No private mailbox data, secrets, or customer identifiers are included.

## What NOT to Do

- Do not commit directly to `main`.
- Do not create releases without PRs.
- Do not claim Outlook behavior is CI-verified until the self-hosted runner exists.
- Do not document deleted presentation tools as active features.
- Do not run OutlookMcp against mailbox data you are not allowed to expose.
- Do not update version numbers manually unless the release process requires it.

## Tips for Good PRs

### Commit Messages

```text
Good: Add duplicate-send guard for mail send
Bad: fix stuff
```

### PR Titles

```text
Good: Add calendar appointment update validation
Bad: Update code
```

### PR Size

- Keep PRs focused.
- Break large changes into smaller chunks.
- Include tests and docs with user-visible changes.

## Local Development Setup

```powershell
git clone https://github.com/trsdn/mcp-server-outlook.git
Set-Location mcp-server-outlook
dotnet restore
dotnet build --nologo -v q
.\src\OutlookMcp.CLI\bin\Debug\net10.0-windows\outlookcli.exe --version
```

To test real Outlook behavior, start classic Outlook first, then run a command such as:

```powershell
.\src\OutlookMcp.CLI\bin\Debug\net10.0-windows\outlookcli.exe application get-status
```

## Trimming and Native AOT Compatibility

### Why Trimming Is Not Supported

OutlookMcp cannot currently be trimmed or published as Native AOT because it depends on Office COM interop and dynamic runtime activation.

### Technical Constraints

Runtime COM activation:

```csharp
Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
dynamic outlook = Activator.CreateInstance(outlookType)!;
```

The trimmer cannot know:

- Which COM type the Windows registry will resolve.
- Which members late-bound COM calls will access.
- Which Office interop assemblies and COM metadata are required at runtime.

### Suppressed Warnings

Warnings such as Windows-only platform checks and trim/AOT incompatibility warnings may be suppressed centrally when they are inherent to Office COM automation.

### Alternatives for Smaller Binaries

- Use framework-dependent deployment when possible.
- Use self-contained deployment only when the target machine may not have the required .NET runtime.

## Need Help?

- Read the docs: [Contributing Guide](CONTRIBUTING.md)
- Ask questions: create a GitHub issue with the `question` label.
- Report bugs: use the bug report template.

---

Every change, no matter how small, must go through a pull request.
