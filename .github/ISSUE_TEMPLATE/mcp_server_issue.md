---
name: MCP Server Issue
about: Report issues with the MCP Server for AI assistants
title: '[MCP] '
labels: 'mcp-server'
assignees: ''

---

## Check this first

Ask your assistant to call `application.get-status` and paste the result.

If it reports `NewOutlookOnly`, this is **not a bug**. The new Outlook for Windows exposes no COM
object model and cannot be automated. Install or switch to the classic Outlook for Windows desktop
app.

```json
[application.get-status result here]
```

## Issue Description
A clear and concise description of the MCP Server issue.

## AI Assistant
Which AI assistant are you using with the MCP Server?
- [ ] **GitHub Copilot** (VS Code, Visual Studio)
- [ ] **Claude Desktop** (Anthropic)
- [ ] **Other**: [please specify]

## MCP Tool & Action
- **Tool**: [one of: mail, calendar, folder, attachment, application]
- **Action**: [e.g. list, search, read, create-draft, send, save]
- **Parameters**: [describe the parameters used, with mailbox content redacted]

## Expected Behavior
What did you expect the MCP Server to do?

## Actual Behavior
What did it actually do?

## Error Response
Paste the full JSON response. It includes `success`, `errorMessage`, and `suggestedNextActions`;
please include all three.

```json
{
  "success": false,
  "errorMessage": "...",
  "suggestedNextActions": []
}
```

## MCP Server Configuration

**Configuration file location**: [e.g. `%APPDATA%\Claude\claude_desktop_config.json`]

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "outlook-mcp-server.exe",
      "args": []
    }
  }
}
```

## Environment
- **Windows Version**: [e.g. Windows 11, Windows 10]
- **Outlook**: classic Outlook for Windows [e.g. Microsoft 365, Outlook 2021, Outlook 2019]
- **OutlookMcp Version**: [e.g. v1.0.0]
- **.NET Version**: [run `dotnet --version`]
- **Installation Method**:
  - [ ] MCPB bundle
  - [ ] VS Code extension
  - [ ] Global .NET tool
  - [ ] Source build
  - [ ] Other: [please specify]
- **Mailbox type**: [Exchange / Microsoft 365 / IMAP / POP / local PST]

## MCP Server Logs
```
[Paste logs here, with mailbox content and addresses redacted]
```

## Steps to Reproduce
1. Configure the AI assistant with the MCP Server
2. Ask the assistant: "..."
3. The server receives a request for tool [tool], action [action]
4. See error

## Before Reporting

- [ ] Classic Outlook was **installed and already running**
- [ ] Outlook had no modal dialog open (a dialog blocks COM calls)
- [ ] Outlook and the MCP server run as the **same Windows user** and at the same elevation
- [ ] If the failure involves an entry ID, it came from a fresh `list`, `search`, or `read-active`.
      Entry IDs change when an item moves between stores.

## Conversation Context (Optional)
The exchange that led to the issue, with content redacted:

```
User: "Show me my unread mail from this week"
AI: [response]
[MCP Server error occurs]
```

## Additional Context

Please do **not** attach real mail items, `.pst`/`.ost` files, screenshots of your inbox, or
anything containing personal data. Describe the shape of the data instead.
