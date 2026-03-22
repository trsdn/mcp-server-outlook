# Outlook MCP Server Skill

Agent Skill for AI assistants using the Outlook MCP Server via MCP.

## Best For

- VS Code Chat and other conversational MCP clients
- Mailbox discovery, draft creation, and attachment export
- Safe reply / reply-all / forward workflows
- Explicit move / delete / read-state workflows after inspection

## Installation

The `Outlook MCP Server (Migration)` VS Code extension bundles this skill automatically.

For manual installs, place this folder in your assistant's skills directory as `outlook-mcp/`.

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

This is the active MCP skill shipped by the Outlook migration surface.

## Related

- `skills/outlook-cli/` for CLI-oriented coding-agent workflows
- `https://github.com/trsdn/mcp-server-outlook`
