# outlook-mcp-skill

An [Agent Skill](https://agentskills.io) for automating classic Outlook for Windows via the
[Outlook MCP Server](https://github.com/trsdn/mcp-server-outlook).

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent
how to drive Outlook through 5 MCP tools and 30 operations:

- **Mail** (16) - read the active item, read, list, search, create drafts, reply, reply-all,
  forward, send, move, delete, set read state, set categories, set subject, body, and recipients
- **Calendar** (5) - list and read appointments, create, update, and delete them
- **Folders** (4) - list default folders, list children, resolve a path, list items
- **Attachments** (4) - list, save, add, remove
- **Application** (1) - report Outlook availability

It also carries behavioural guidance: confirm before sending or deleting, discover rather than ask,
and address items by entry ID.

## Requirements

- Windows
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows exposes no COM object model and cannot be automated; `application.get-status` reports
  `NewOutlookOnly` in that case.
- The [Outlook MCP Server](https://github.com/trsdn/mcp-server-outlook) configured in your client

## Install

```powershell
npx skillpm install outlook-mcp-skill
```

Or with npm directly:

```powershell
npm install outlook-mcp-skill
```

## License

MIT
