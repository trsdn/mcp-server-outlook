# Installation Guide - OutlookMcp

Complete installation instructions for the OutlookMcp MCP server and `outlookcli` CLI.

## System Requirements

### Required

- Windows 10 or later.
- Classic Outlook for Windows desktop app installed and running.
- .NET 10 Runtime or SDK for .NET tool installs. The VS Code extension and MCPB package are self-contained when published.

OutlookMcp automates Outlook through the `Outlook.Application` COM ProgID. The new Outlook for Windows does not expose that COM object model and cannot be automated by this project. If only new Outlook is installed, `application get-status` reports `NewOutlookOnly`.

### Optional

- Node.js, only for `npx` helpers such as `add-mcp` or `skills` installation.
- Visual Studio 2022 or VS Code, only for source development.

### Recommended

- Windows 11 with the latest Microsoft 365 Apps updates.
- Start classic Outlook before invoking MCP or CLI commands.
- Run OutlookMcp at the same elevation level as Outlook. Do not run one elevated and the other unelevated.

---

## Quick Start (Recommended)

Use this order to avoid setup confusion:

1. Choose one primary MCP setup path:
   - VS Code extension for GitHub Copilot users.
   - Claude Desktop MCPB package, when available from releases.
   - Manual MCP setup for other MCP clients.
2. Start classic Outlook for Windows.
3. Validate with `application get-status`.
4. Optionally install `outlookcli` for scripting.
5. Optionally install agent skills for clients that support skills.

### VS Code Extension

**Best for:** GitHub Copilot users who want automatic MCP configuration.

1. Open VS Code.
2. Press `Ctrl+Shift+X`.
3. Search for `OutlookMcp`.
4. Install the extension.
5. Restart VS Code if prompted.

The extension is expected to configure the MCP server for GitHub Copilot and include the packaged server bits. If the marketplace package has not been updated for the current release, use the manual MCP setup below.

**Marketplace Link:** [Outlook MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=trsdn.outlook-mcp)

---

### Claude Desktop MCPB

**Best for:** Claude Desktop users who want a packaged install.

1. Download `outlook-mcp-{version}.mcpb` from the [latest release](https://github.com/trsdn/mcp-server-outlook/releases/latest), if available.
2. Double-click the `.mcpb` file or drag it onto Claude Desktop.
3. Restart Claude Desktop.
4. Start classic Outlook and validate with the prompt below.

---

## Manual MCP Setup (All MCP Clients)

**Best for:** Cursor, Windsurf, Cline, Claude Code, Codex, and advanced users.

### Step 1: Install .NET 10

Check if it is already installed:

```powershell
dotnet --version
```

Install the runtime if needed:

```powershell
winget install Microsoft.DotNet.Runtime.10
```

Manual download: [.NET 10 Downloads](https://dotnet.microsoft.com/download/dotnet/10.0)

### Step 2: Install OutlookMcp MCP Server

```powershell
dotnet tool install --global OutlookMcp.McpServer
dotnet tool list --global | Select-String "OutlookMcp"
```

Optional CLI install:

```powershell
dotnet tool install --global OutlookMcp.CLI
outlookcli --version
```

If you install both tools, keep them on the same version.

### Step 3: Configure Your MCP Client

#### Option A: Auto-Configure All Agents

Use `add-mcp` to configure detected clients:

```powershell
npx add-mcp "mcp-outlook" --name outlook-mcp
```

Examples:

```powershell
npx add-mcp "mcp-outlook" --name outlook-mcp -a cursor -a claude-code
npx add-mcp "mcp-outlook" --name outlook-mcp -g
npx add-mcp "mcp-outlook" --name outlook-mcp --all -y
```

Requires Node.js for `npx`:

```powershell
winget install OpenJS.NodeJS.LTS
```

#### Option B: Manual Configuration

Ready-to-use config files are in [`examples/mcp-configs/`](https://github.com/trsdn/mcp-server-outlook/tree/master/examples/mcp-configs/).

**For GitHub Copilot in VS Code**, create `.vscode\mcp.json`:

```json
{
  "servers": {
    "outlook-mcp": {
      "command": "mcp-outlook"
    }
  }
}
```

**For GitHub Copilot in Visual Studio**, create `.mcp.json` in your solution directory or `%USERPROFILE%\.mcp.json`:

```json
{
  "servers": {
    "outlook-mcp": {
      "command": "mcp-outlook"
    }
  }
}
```

**For Claude Desktop**, merge this into `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "mcp-outlook",
      "args": [],
      "env": {}
    }
  }
}
```

**For Cursor, Cline, or Windsurf**, add the same `outlook-mcp` server entry to that client's MCP configuration file.

### Step 4: Validate MCP Setup

Start classic Outlook, restart your MCP client, then ask:

```text
Use OutlookMcp to get Outlook application status.
```

Expected result: the `application` tool with `get-status` reports classic Outlook availability. If it reports `NewOutlookOnly`, switch to or install classic Outlook for Windows.

---

## Optional: CLI Installation (No AI Required)

**Best for:** Scripting, RPA, and direct automation.

```powershell
dotnet tool install --global OutlookMcp.CLI
outlookcli --version
```

### Quick Tests

```powershell
outlookcli application get-status
outlookcli folder list-default
outlookcli mail list --folder Inbox --max-count 10 --include-body-preview
```

The current Outlook command surface is sessionless. Mail, calendar, folder, attachment, and application commands run against the user's already-running classic Outlook desktop app.

Use `--quiet` or `-q` to suppress the banner for scripting:

```powershell
outlookcli -q mail list --folder Inbox --max-count 5
```

CLI documentation: [CLI Guide](https://github.com/trsdn/mcp-server-outlook/blob/master/src/OutlookMcp.CLI/README.md)

---

## Agent Skills Installation

**Best for:** Adding tool-specific guidance to coding agents that support skills.

Skills are installed by the VS Code extension when that package is current. For other clients:

```powershell
npx skills add trsdn/mcp-server-outlook --skill outlook-cli
npx skills add trsdn/mcp-server-outlook --skill outlook-mcp
npx skills add trsdn/mcp-server-outlook --skill outlook-cli -a cursor
npx skills add trsdn/mcp-server-outlook --skill outlook-mcp -a claude-code
npx skills add trsdn/mcp-server-outlook --skill outlook-cli --global
```

See [Agent Skills Guide](../skills/README.md).

---

## Updating OutlookMcp

### Check Installed Version

```powershell
dotnet tool list --global | Select-String "OutlookMcp"
outlookcli --version
mcp-outlook --version
```

### Update Installed Tools

```powershell
dotnet tool update --global OutlookMcp.McpServer
dotnet tool update --global OutlookMcp.CLI
```

Restart your MCP client after updating the MCP server.

### Troubleshooting Updates

#### Update Command Fails

```powershell
dotnet tool uninstall --global OutlookMcp.McpServer
dotnet tool install --global OutlookMcp.McpServer
```

#### MCP Server Still Running Old Version

Close and reopen the MCP client. Some clients keep server processes alive until the host application exits.

### Rollback to Previous Version

```powershell
dotnet tool uninstall --global OutlookMcp.McpServer
dotnet tool install --global OutlookMcp.McpServer --version 1.2.3
```

Replace `1.2.3` with the version you need.

### Check What Changed

- GitHub Releases: <https://github.com/trsdn/mcp-server-outlook/releases>
- Changelog: <https://github.com/trsdn/mcp-server-outlook/blob/master/CHANGELOG.md>

---

## Troubleshooting

### Common Issues

#### 1. `dotnet` command not found

Install .NET 10 Runtime or SDK.

#### 2. MCP Server Not Responding

```powershell
dotnet tool list --global | Select-String "OutlookMcp"
dotnet tool uninstall --global OutlookMcp.McpServer
dotnet tool install --global OutlookMcp.McpServer
```

Then restart the MCP client.

#### 3. Classic Outlook Not Running

Start classic Outlook for Windows, sign in, let it finish loading, and retry.

#### 4. New Outlook Only

New Outlook for Windows has no COM object model. Install or switch to classic Outlook for Windows.

#### 5. Elevation Mismatch

If Outlook runs normally but the server or CLI runs as Administrator, COM may not find the running Outlook instance. Run both at the same elevation level.

#### 6. Entry ID Not Found After Moving an Item

Outlook entry IDs can change when an item moves between stores. Re-run list, search, or read-active and use the new entry ID.

## Uninstallation

### Uninstall MCP Server

```powershell
dotnet tool uninstall --global OutlookMcp.McpServer
```

### Uninstall CLI

```powershell
dotnet tool uninstall --global OutlookMcp.CLI
```

---

## Getting Help

- Documentation: [GitHub Repository](https://github.com/trsdn/mcp-server-outlook)
- Issues: [GitHub Issues](https://github.com/trsdn/mcp-server-outlook/issues)
- Contributing: [Contributing Guide](https://github.com/trsdn/mcp-server-outlook/blob/master/docs/CONTRIBUTING.md)

---

## Next Steps

After installation:

1. Validate status with `application get-status`.
2. Explore folders with `folder list-default`.
3. Read mail using `mail list`, `mail search`, and `mail read-active`.
4. Use entry IDs from list, search, or read-active results for follow-up operations.
5. Remember that OutlookMcp acts on the real Outlook mailbox in the shared desktop app.

---

Happy automating!
