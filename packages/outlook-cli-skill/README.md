# outlook-cli-skill

An [Agent Skill](https://agentskills.io) for automating classic Outlook for Windows via the
[`outlookcli`](https://github.com/trsdn/mcp-server-outlook) command-line tool.

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent
how to drive Outlook from scripts, with the same 8 tools and 66 operations the MCP server exposes:

- **Mail** (23) - read the active item, read, list, search, read a whole conversation, respond to a
  meeting invitation, create drafts, reply, reply-all, forward, send, move, export, delete, set read
  state, flags, categories, subject, body and recipients, and list categories, rules and reminders
- **Calendar** (7) - list and read appointments, create, update and delete them, check free/busy, export
- **Contacts** (5) - list and read contacts, create, update and delete them
- **Tasks** (5) - list and read tasks, create, update and delete them
- **Rules** (5) - list the inbox rules that decide what happens to mail before you read it, and
  create, change, switch off or remove them
- **Folders** (10) - list default folders and stores, open a shared mailbox, list children, resolve
  a path, list items, and create, rename, move or delete folders
- **Attachments** (4) - list, save, add, remove
- **Application** (3) - report Outlook availability and the active explorer or inspector

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