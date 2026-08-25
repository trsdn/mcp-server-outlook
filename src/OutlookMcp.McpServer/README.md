# Outlook MCP server migration surface

This MCP server area is the active MCP host surface for `mcp-server-outlook`.

Current reality:

- the repository target is now `mcp-server-outlook`
- the public MCP surface exposes a working Outlook seed for application, attachment, folder, and mail workflows
- the internal service and generator pipeline are reusable, but the broader tool taxonomy still needs more Outlook-first families
- the current tool list is an early Outlook contract, not the final one

## Implemented Outlook seed

The generated MCP surface now includes these Outlook-first categories:

- `application`
- `attachment`
- `calendar`
- `folder`
- `mail`

Current seed actions:

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

## What will remain useful

- host/service split
- named-pipe communication
- generator-based MCP registration
- result models and error handling patterns
- packaging and distribution pipeline

## What still has to change

- PowerPoint-only tool families such as `slide`, `shape`, `animation`, `transition`, and `slideshow`
- file-centric workflow assumptions
- schema descriptions, examples, and install guidance that still describe PowerPoint behavior

## Intended next Outlook-first MCP families

The next Outlook-first MCP expansion should center on:

- `application`
- `session`
- `folder`
- `mail`
- `attachment`
- `calendar`
- `contact`

## Naming note

Until the cleanup pass lands, expect inherited names nearby such as `OutlookMcp.*`, `mcp-outlook`, and `outlook-mcp`. Those are transitional and not the intended long-term Outlook branding.
