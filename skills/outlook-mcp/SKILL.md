---
name: outlook-mcp
description: >
  Automate Microsoft Outlook on Windows via MCP and COM interop. Use when inspecting Outlook
  application state, listing folders, reading active mail, searching mail, creating drafts,
  replying safely, moving or deleting mail intentionally, updating read state, sending
  explicitly, or exporting attachments.
  Triggers: Outlook, email, mailbox, draft, reply, attachment, send, folder.
---

# Outlook MCP Server Skill

Agent Skill for AI assistants using the Outlook MCP Server via the Model Context Protocol.

## Best For

- Conversational AI in VS Code Chat and similar MCP-capable clients
- Iterative mailbox exploration with rich tool schemas
- Safe draft-first workflows before explicit send actions
- Attachment discovery and export workflows

## Recommended workflow

1. Start with `application.get-status` to confirm Outlook availability.
2. Use `folder.list-default` to discover useful mailbox anchors like Inbox and Drafts.
3. Use `mail.list`, `mail.search`, or `mail.read-active` to inspect current context.
4. Create or derive drafts with `mail.create-draft`, `mail.reply`, `mail.reply-all`, or `mail.forward`.
5. Use `mail.set-read-state`, `mail.set-categories`, `mail.move`, or `mail.delete` only after you have identified the exact message.
6. Use `calendar.list` or `calendar.read` before creating a new appointment if the task depends on existing schedule context.
7. Use `attachment.list` and `attachment.save` when files are part of the task.
8. Only use `mail.send` as an explicit final action when the draft and recipients are already correct.

## Safety rules

- Prefer draft-producing actions over immediate mailbox mutation.
- Treat `mail.send` as explicit and intentional.
- If an item is already selected in Outlook, inspect it first with `mail.read-active` before changing anything.
- Treat `mail.move` and `mail.delete` as explicit mailbox mutations after confirmation by context.
- Prefer `attachment.list` before `attachment.save` so you know exactly what will be exported.

## Current Outlook seed

- `application.get-status`
- `folder.list-default`
- `folder.list-children`
- `mail.read-active`
- `mail.read`
- `mail.list`
- `mail.search`
- `mail.create-draft`
- `mail.reply`
- `mail.reply-all`
- `mail.forward`
- `mail.send`
- `mail.move`
- `mail.delete`
- `mail.set-read-state`
- `mail.set-categories`
- `calendar.list`
- `calendar.read`
- `calendar.create-appointment`
- `calendar.update-appointment`
- `calendar.delete-appointment`
- `attachment.list`
- `attachment.add`
- `attachment.remove`
- `attachment.save`

## Transitional naming note

The repository is still migrating from inherited `OutlookMcp.*` internals. Public guidance should be Outlook-first even when executable or project names still contain `OutlookMcp`.
