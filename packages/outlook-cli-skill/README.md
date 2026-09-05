# outlook-cli-skill

An [Agent Skill](https://agentskills.io) for automating classic Outlook for Windows via the
[`outlookcli`](https://github.com/trsdn/mcp-server-outlook) command-line tool.

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent
how to drive Outlook from scripts, with the same 10 tools and 69 operations the MCP server exposes:

- **Mail** (23) - read the active item, read, list, search, read a whole conversation, answer a
  meeting invitation, create drafts, reply, reply-all, forward, send, move, delete, export, and set
  read state, flags, categories, subject, body and recipients
- **Folders** (10) - list default folders and stores, open a shared mailbox, create, rename, move
  and delete folders, list children, resolve a path, list items
- **Calendar** (7) - list and read appointments, create, update and delete them, check free/busy,
  export
- **Contacts** (5) - list, read, create, update, delete
- **Tasks** (5) - list, read, create, update, delete
- **Attachments** (4) - list, save, add, remove
- **Application** (3) - report Outlook availability and what the user is looking at
- **Address book** (3) - resolve addressees to real SMTP addresses before sending, list address
  books, browse one
- **Message properties** (4) - read internet message headers, MAPI properties and custom user
  properties

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
