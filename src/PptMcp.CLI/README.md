# Outlook CLI migration surface

This CLI area is the current command-line migration surface for `mcp-server-outlook`.

Current reality:

- the repo target is now Outlook-first
- the generated command surface includes a working Outlook seed for application, folder, mail, and attachment workflows
- the single-entry CLI pattern remains useful for coding agents while the broader Outlook taxonomy is still being expanded

## Implemented Outlook seed

The generated CLI/MCP command model now includes these first Outlook categories:

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

## Intended Outlook-first CLI workflow

The Outlook CLI should evolve toward flows such as:

- connect to Outlook / establish working context
- list folders or set current folder
- inspect and export attachments
- list, search, and inspect mail items
- create or edit drafts
- reply, reply-all, forward, send
- manage attachments
- create and inspect appointments
- inspect and update contacts

## What is still transitional right now

- inherited `pptcli` naming in some areas
- PowerPoint examples in help and docs
- generator output based on PowerPoint interfaces

Those are migration leftovers around the CLI shell, not the intended Outlook product shape.
