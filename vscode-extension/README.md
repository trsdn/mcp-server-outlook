# mcp-server-outlook VS Code Extension - Migration Status

This extension area is the active VS Code packaging surface for the Outlook migration.

Current reality:

- the repository target is now `mcp-server-outlook`
- the published extension now exposes Outlook-first provider ids and bundled Outlook skill folders
- the bundled binaries still use inherited `OutlookMcp.*` executable names during migration
- the extension now documents the real Outlook seed instead of the legacy PowerPoint product

## Implemented Outlook seed

The underlying generated server surface now includes an initial Outlook seed:

- `application.get-status`
- `attachment.list`
- `attachment.add`
- `attachment.remove`
- `attachment.save`
- `calendar.list`
- `calendar.read`
- `calendar.create-appointment`
- `calendar.update-appointment`
- `calendar.delete-appointment`
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

The extension now surfaces the Outlook seed as the primary story, while deeper Outlook families still need follow-up work.

## What this area needs to become

The VS Code extension should become an Outlook-first marketplace offering for workflows such as:

- email triage and drafting
- reply / reply-all / forward flows
- attachment inspection and export
- attachment handling
- calendar scheduling and meeting updates
- mailbox and folder navigation

## What is still in progress

- package and identifier rename
- skill follow-through across all remaining legacy docs and references
- marketplace rebrand
- cleanup of inherited PowerPoint examples and help text
- wiring the extension to real Outlook-first MCP tool families

## Important naming note

Until the cleanup pass lands, expect inherited names in adjacent files and packaging metadata, including:

- `OutlookMcp.*` project names
- `outlook-mcp` extension/package identifiers
- `mcp-outlook` and `outlookcli` command names in some docs and configs

Those names are transitional, not the intended long-term Outlook branding.
