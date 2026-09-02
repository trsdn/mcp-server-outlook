# MCP Configuration Examples

This directory contains ready-to-use MCP configuration files for AI coding assistants and desktop clients.

## Quick Setup Guide

### 1. Install OutlookMcp MCP Server

```powershell
dotnet tool install --global OutlookMcp.McpServer
```

Classic Outlook for Windows must be installed and running. New Outlook for Windows is not supported because it has no `Outlook.Application` COM object model.

### 2. Choose Your Client and Copy the Config

Select the configuration file for your AI assistant and follow the instructions below.

---

## Claude Desktop

**Config File:** `claude-desktop-config.json`

**Location:** `%APPDATA%\Claude\claude_desktop_config.json` on Windows

**Setup Steps:**

1. Open File Explorer and navigate to `%APPDATA%\Claude\`.
2. If `claude_desktop_config.json` does not exist, create it.
3. Copy the contents of `claude-desktop-config.json` from this folder.
4. If you already have a config file, merge the `outlook-mcp` server entry into your existing `mcpServers` section.
5. Restart Claude Desktop.
6. Start classic Outlook.

**Test it:**

```text
Use OutlookMcp to get Outlook application status.
```

---

## Cursor

**Config File:** `cursor-mcp-config.json`

**Location:**

- Windows: `%APPDATA%\Cursor\User\globalStorage\mcp\mcp.json`
- Project-specific: `.cursor\mcp.json` in your workspace

**Setup Steps:**

1. Open Cursor Settings with `Ctrl+,`.
2. Search for "MCP".
3. Click "Edit in settings.json" or manually create the config file.
4. Copy the contents of `cursor-mcp-config.json` from this folder.
5. If you already have a config file, merge the `outlook-mcp` server entry.
6. Restart Cursor.
7. Start classic Outlook.

**Test it:**

```text
Use OutlookMcp to list my default Outlook folders.
```

---

## Cline (VS Code Extension)

**Config File:** `cline-mcp-config.json`

**Location:**

- VS Code user settings through the MCP settings icon in Cline.
- Or manually: `%APPDATA%\Code\User\globalStorage\saoudrizwan.claude-dev\settings\cline_mcp_settings.json` on Windows.

**Setup Steps:**

1. Install the Cline extension in VS Code.
2. Open the Cline panel.
3. Click the MCP settings gear icon.
4. Add the server configuration from `cline-mcp-config.json`.
5. Restart VS Code.
6. Start classic Outlook.

**Test it:**

```text
Use OutlookMcp to get Outlook application status.
```

---

## Windsurf

**Config File:** `windsurf-mcp-config.json`

**Location:**

- Windows: `%APPDATA%\Windsurf\User\mcp_settings.json`
- Or check Windsurf's MCP settings panel.

**Setup Steps:**

1. Open Windsurf Settings.
2. Navigate to MCP Servers configuration.
3. Add the server configuration from `windsurf-mcp-config.json`.
4. Restart Windsurf.
5. Start classic Outlook.

**Test it:**

```text
Use OutlookMcp to list my default Outlook folders.
```

---

## VS Code (GitHub Copilot)

**Config File:** `vscode-mcp-config.json`

**Location:** `.vscode\mcp.json` in your workspace

**Setup Steps:**

**Option A: Use VS Code Extension**

1. Install the [Outlook MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=trsdn.outlook-mcp).
2. Configuration is automatic when the extension package is current.
3. Start classic Outlook.

**Option B: Manual Configuration**

1. Create `.vscode\mcp.json` in your project.
2. Copy contents from `vscode-mcp-config.json`.
3. Reload the VS Code window.
4. Start classic Outlook.

**Test it:**

```text
Use OutlookMcp to get Outlook application status.
```

---

## Troubleshooting

### Server Not Responding

1. Verify installation:

```powershell
dotnet tool list --global | Select-String "OutlookMcp"
```

2. Check .NET:

```powershell
dotnet --version
```

3. Reinstall if needed:

```powershell
dotnet tool uninstall --global OutlookMcp.McpServer
dotnet tool install --global OutlookMcp.McpServer
```

4. Restart the MCP client.

### Classic Outlook Not Found

- Ensure classic Outlook for Windows is installed.
- Start Outlook and wait until it is signed in and usable.
- Do not use new Outlook for Windows.
- Run Outlook and the MCP server at the same elevation level.

### Permission or Data Issues

- OutlookMcp acts on the real mailbox in the running Outlook desktop app.
- Review destructive operations before allowing them.
- Entry IDs can change when items move between stores. Re-list or search when an item cannot be found.

### Still Having Issues?

- Check the [main installation guide](../../docs/INSTALLATION.md).
- Report issues on [GitHub](https://github.com/trsdn/mcp-server-outlook/issues).

---

## Configuration Options

### Multiple Workspaces

If you work with multiple workspaces, you can:

- Use project-specific config files when the client supports them.
- Use global user-level configuration for all projects.

---

## Learn More

- [Main README](../../README.md) - Feature overview and examples.
- [Installation Guide](../../docs/INSTALLATION.md) - Comprehensive setup instructions.
- [MCP Server README](../../src/OutlookMcp.McpServer/README.md) - Tool documentation.
- [GitHub Repository](https://github.com/trsdn/mcp-server-outlook) - Source code and issues.
