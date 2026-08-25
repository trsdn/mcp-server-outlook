# Outlook CLI Skill

Agent Skill for coding assistants using the inherited Outlook CLI surface.

## Best For

- Coding agents that prefer CLI commands over MCP tools
- Scriptable mailbox inspection and draft workflows
- Quiet, deterministic automation around Outlook seed actions
- Explicit mailbox mutation flows once the target mail item is identified

## Transitional note

The CLI binary naming is still inherited from the copied PowerPoint repo, so some commands and package names may still reference `OutlookMcp` or `outlookcli` during migration.

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

This is the active CLI skill for the current Outlook migration surface.
