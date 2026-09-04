# mcp-server-outlook VS Code Extension

The VS Code packaging surface for the Outlook MCP server.

## What it provides

The extension bundles the Outlook MCP server and registers it with VS Code, so Copilot Chat and
other MCP clients can drive the classic Outlook desktop app.

The server exposes **$18 tools with 62 operations**:

| Tool | Operations |
|---|---|
| `mail` | `read-active`, `read`, `list`, `search`, `create-draft`, `reply`, `reply-all`, `forward`, `send`, `move`, `delete`, `set-read-state`, `set-categories`, `set-subject`, `set-body`, `set-recipients` |
| `calendar` | `list`, `read`, `create-appointment`, `update-appointment`, `delete-appointment` |
| `folder` | `list-default`, `list-children`, `resolve-path`, `list-items` |
| `attachment` | `list`, `save`, `add`, `remove` |
| `application` | `get-status` |

The CLI (`outlookcli`) exposes exactly the same actions with the same parameters. See
[FEATURES.md](../FEATURES.md).

## Requirements

- Windows
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows has no COM object model and cannot be automated; in that case `application.get-status`
  reports `NewOutlookOnly` and every action fails with an actionable message.

Run `application.get-status` first to confirm the environment before anything else.

## What is still in progress

- package and identifier rename: the bundled binaries still use inherited `OutlookMcp.*` executable
  names, and the extension id is still transitional
- marketplace rebrand
- richer server-side mail search (#42) and a paging cursor for large result sets (#43)
- **no Outlook behaviour is verified by CI.** Integration tests need a self-hosted Windows runner
  with classic Outlook installed, which does not exist yet (#31)

## Naming note

Expect inherited names in adjacent files and packaging metadata:

- `OutlookMcp.*` project and assembly names
- `outlook-mcp` extension and package identifiers
- `outlookcli` command name

These are tracked under #12. They are transitional, not the intended long-term branding.
