# outlook-mcp-skill

An [Agent Skill](https://agentskills.io) for automating classic Outlook for Windows via the
[Outlook MCP Server](https://github.com/trsdn/mcp-server-outlook).

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent
how to drive Outlook through 8 MCP tools and 66 operations:

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