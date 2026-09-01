---
name: outlook-cli
description: >
  Automate Microsoft Outlook on Windows via CLI. Use when coding agents prefer short commands
  over MCP schemas for mailbox inspection, draft creation, mail-state changes, sending, or
  attachment export.
  Triggers: Outlook, email, mailbox, draft, attachment, send, outlookcli, CLI automation.
---

# Outlook automation with the CLI

## Preconditions

- Windows host with Microsoft Outlook installed
- The current migration still uses inherited `OutlookMcp.CLI` / `outlookcli` naming
- Use the generated CLI help to confirm exact flags while the Outlook surface expands

## Recommended workflow

1. Discover available Outlook categories in CLI help.
2. Check Outlook status with the `application` category.
3. Inspect mailbox context with `folder` and `mail`.
4. Create safe drafts first, then mutate mailbox state only after identifying the exact item.
5. Use the `calendar` category for schedule inspection and safe appointment creation.
6. Send explicitly only when ready.
7. Export attachments with the `attachment` category.

## Current Outlook seed

- `application.get-status`
- `folder.list-default`
- `folder.list-children`
- `folder.resolve-path`
- `folder.list-items`
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
- `mail.set-subject`
- `mail.set-body`
- `mail.set-recipients`
- `calendar.list`
- `calendar.read`
- `calendar.create-appointment`
- `calendar.update-appointment`
- `calendar.delete-appointment`
- `attachment.list`
- `attachment.add`
- `attachment.remove`
- `attachment.save`

## Safety rules

- Prefer inspection before mutation.
- Prefer draft creation before send.
- Treat `mail.move`, `mail.delete`, `mail.set-read-state`, and `mail.set-categories` as explicit mailbox mutations.
- Treat attachment export paths as explicit user-controlled destinations.
- End with a short text summary after CLI execution.
