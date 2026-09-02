# Claude Desktop Configuration

The Outlook MCP Server works with Claude Desktop on Windows. It automates the **classic Outlook for
Windows desktop app** through COM, so that app must be installed and running on the same machine.

## Configuration Location

```
%APPDATA%\Claude\claude_desktop_config.json
```

## Basic Configuration

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

Or using the .NET tool:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "dotnet",
      "args": ["outlook-mcp-server"]
    }
  }
}
```

## Windows Considerations

### Outlook instance

- The server attaches to the Outlook instance the user is already running, rather than starting its
  own hidden copy. Operations act on the real mailbox.
- There is no window show/hide control. Outlook stays as the user left it.
- A modal dialog open in Outlook can block COM calls until it is dismissed.

### File system access

Claude Desktop runs with limited file system access, which matters for `attachment.save` and
`attachment.add`. Prefer paths under:

- `C:\Users\<username>\Documents\`
- `C:\Users\<username>\Desktop\`
- `%TEMP%`

### Session persistence

The server holds no document session. Outlook itself is the state, so closing Claude Desktop does
not discard work. Drafts you created remain in the Drafts folder.

## Recommended First Call

```
application.get-status
```

This reports whether classic Outlook is available. If it returns `NewOutlookOnly`, stop: the new
Outlook for Windows exposes no COM object model and cannot be automated.

## Troubleshooting

### `NewOutlookOnly` status

Only the new Outlook for Windows is installed. Install classic Outlook, or switch the "New Outlook"
toggle off, then restart the server.

### "Outlook not found" / COM activation failure

- Confirm classic Outlook is installed and has been launched at least once.
- Confirm the `Outlook.Application` ProgID is registered.
- Claude Desktop and Outlook must run as the same Windows user. An elevated Outlook and a
  non-elevated server will not connect.

### "Access denied" on attachment paths

Use a path under Documents, Desktop, or `%TEMP%`.

### Calls hang or time out

- Outlook is probably showing a dialog. Check for a modal window and dismiss it.
- A very large `mail.list` or `mail.search` over a busy mailbox is slow. Constrain by folder and
  count.

### An entry ID stopped resolving

Entry IDs change when an item moves between stores. Re-run the `list` or `search` that produced it.

## MCPB Bundle Alternative

For simplified installation, use the MCPB bundle, which auto-configures Claude Desktop:

1. Download `outlook-mcp-bundle.mcpb` from releases
2. Double-click to install
3. Restart Claude Desktop

See [`mcpb/README.md`](../../../mcpb/README.md) for details.
