# Examples

## MCP client configuration

[`mcp-configs/`](mcp-configs/) holds ready-to-use MCP server configuration snippets for:

- Claude Desktop
- VS Code
- Cursor
- Cline
- Windsurf

Copy the relevant file's contents into your client's MCP configuration.

## Requirements

- Windows
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows has no COM object model and cannot be automated.

Run `application.get-status` (or `outlookcli application get-status`) first to confirm your
environment before anything else.

## CLI usage

The CLI exposes the same 9 tools and 61 actions as the MCP server, with the same parameters:

```powershell
outlookcli application get-status
outlookcli folder list-default
outlookcli mail list --folder Inbox
outlookcli attachment list --entry-id <id>
```

Use `--help` on any tool or action to see its parameters:

```powershell
outlookcli mail --help
outlookcli mail search --help
```

See [FEATURES.md](../FEATURES.md) for the full action list.

## A note on sessions

There is no session or batch concept in the Outlook surface. Outlook is a shared desktop
application the server attaches to, not a document it opens and closes, so there is nothing to open,
save, or dispose. Each call acts on the live mailbox immediately.

