# outlook-cli-skill

An [Agent Skill](https://agentskills.io) for automating classic Outlook for Windows via the
[`outlookcli`](https://github.com/trsdn/mcp-server-outlook) command-line tool.

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent
how to drive Outlook from scripts, with the same 10 tools and 62 operations the MCP server exposes:

- **Mail** (16) - read the active item, read, list, search, create drafts, reply, reply-all,
  forward, send, move, delete, set read state, set categories, set subject, body, and recipients
- **Calendar** (5) - list and read appointments, create, update, and delete them
- **Folders** (4) - list default folders, list children, resolve a path, list items
- **Attachments** (4) - list, save, add, remove
- **Application** (1) - report Outlook availability

The CLI surface is deliberately token-efficient and fully discoverable through `--help`, which makes
it a better fit than the MCP server for coding agents.

## Requirements

- Windows
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows exposes no COM object model and cannot be automated; `outlookcli application get-status`
  reports `NewOutlookOnly` in that case.
- Install the CLI: `dotnet tool install --global OutlookMcp.CLI`

## Install

```powershell
npx skillpm install outlook-cli-skill
```

Or with npm directly:

```powershell
npm install outlook-cli-skill
```

## License

MIT
