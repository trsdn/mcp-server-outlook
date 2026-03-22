---
name: outlook-cli
description: >
  Automate Microsoft Outlook on Windows via CLI. Use when coding agents prefer short commands
  over MCP schemas for mailbox inspection, draft creation, sending, or attachment export.
  Triggers: Outlook, email, mailbox, draft, attachment, send, pptcli, CLI automation.
---

# Outlook automation with the CLI

## Preconditions

- Windows host with Microsoft Outlook installed
- CLI naming may still be inherited (`PptMcp.CLI`, `pptcli`) while migration continues
- Use CLI help to confirm exact arguments for the current generated Outlook seed

## Recommended workflow

1. Inspect Outlook with the `application` category.
2. Discover folders with `folder`.
3. Inspect or search mail with `mail`.
4. Create drafts first, then send explicitly only when ready.
5. Export attachments with `attachment`.

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

## Safety rules

- Prefer inspection before mutation.
- Prefer draft creation before send.
- Treat attachment export paths as explicit destinations.
- End with a short text summary after command execution.
