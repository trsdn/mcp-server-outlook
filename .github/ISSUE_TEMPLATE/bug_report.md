---
name: Bug Report
about: Create a report to help us improve OutlookMcp
title: '[BUG] '
labels: 'bug'
assignees: ''

---

## Check this first

Run `application.get-status` (or `outlookcli application get-status`) and paste the result.

If it reports `NewOutlookOnly`, this is **not a bug**. The new Outlook for Windows exposes no COM
object model and cannot be automated. Install or switch to the classic Outlook for Windows desktop
app.

```
[application.get-status output here]
```

## Bug Description
A clear and concise description of what the bug is.

## Component
Which component is this bug related to?
- [ ] **MCP Server** (Model Context Protocol server for AI assistants)
- [ ] **CLI** (`outlookcli`)
- [ ] **Core Library** (shared functionality)
- [ ] **VS Code extension**
- [ ] **Not sure**

## Command/Usage

**For CLI:**
```powershell
outlookcli <tool> <action> [options]
```

**For MCP Server:**
- Tool: [one of: mail, calendar, folder, attachment, application]
- Action: [e.g. list, search, send, save]
- Parameters used: [describe what was passed]

## Expected Behavior
A clear and concise description of what you expected to happen.

## Actual Behavior
A clear and concise description of what actually happened.

## Error Message
If applicable, paste the full error message. Results include `success`, `errorMessage`, and
`suggestedNextActions`; please include all three.

```json
[Error output here]
```

## Environment
- **Windows Version**: [e.g. Windows 11, Windows 10]
- **Outlook**: classic Outlook for Windows [e.g. Microsoft 365, Outlook 2021, Outlook 2019]
- **OutlookMcp Version**: [e.g. v1.0.0]
- **.NET Version**: [run `dotnet --version`]
- **Installation Method**: [NuGet tool / binary download / MCPB bundle / VS Code extension / source build]
- **AI Assistant** (if using the MCP Server): [e.g. GitHub Copilot, Claude Desktop]
- **Mailbox type**: [Exchange / Microsoft 365 / IMAP / POP / local PST]

## Steps to Reproduce
1. ...
2. ...
3. See error

## Before Reporting

- [ ] Classic Outlook was **installed and already running** when the command ran
- [ ] Outlook had no modal dialog open (a dialog blocks COM calls)
- [ ] Outlook and the server run as the **same Windows user** and at the same elevation
- [ ] If the failure involves an entry ID, it came from a fresh `list`, `search`, or `read-active`
      in this session. Entry IDs change when an item moves between stores.

## Additional Context

Please do **not** attach real mail items, `.pst`/`.ost` files, or anything containing personal data.
Describe the shape of the item instead (folder, whether it has attachments, roughly how many
recipients).
