---
name: outlook-mcp
description: >
  Automate Microsoft Outlook on Windows via MCP and COM interop. Use when inspecting Outlook
  application state, listing folders, reading active mail, searching mail, creating drafts,
  replying safely, sending explicitly, or exporting attachments.
  Triggers: Outlook, email, mailbox, draft, reply, attachment, send, folder.
---

# Outlook MCP Server Skill

Provides 33 Outlook operations via Model Context Protocol. The MCP server forwards
requests to the shared service layer while the repository continues its migration from inherited
`OutlookMcp.*` internals.

## Recommended workflow

1. `application.get-status`
2. `folder.list-default`
3. `mail.list`, `mail.search`, or `mail.read-active`
4. `mail.create-draft`, `mail.reply`, `mail.reply-all`, or `mail.forward`
5. `attachment.list` / `attachment.save` when files matter
6. `mail.send` only as an explicit final action

## Safety rules

- Prefer draft-producing actions before send.
- Inspect current context before mutating mailbox state.
- Treat attachment export destinations as explicit user-controlled paths.
- End with a short text summary after tool use.

## Current Outlook seed

- `application.get-status`
- `folder.list-default`
- `mail.read-active`
- `mail.list`
- `mail.search`
- `mail.create-draft`
- `mail.reply`
- `mail.reply-all`
- `mail.forward`
- `mail.send`
- `attachment.list`
- `attachment.save`

## Transitional note

The repo was renamed from an Office-automation baseline, so some executables or project names
may still contain `OutlookMcp`. Public guidance should nevertheless stay Outlook-first.
