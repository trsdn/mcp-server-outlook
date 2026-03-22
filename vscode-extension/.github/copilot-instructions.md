# Outlook MCP Server - Quick Reference

> When the user asks about Outlook mailboxes, folders, draft mail, replies, sends, or attachments, use the Outlook MCP tools exposed by this extension.

## When to use the Outlook MCP extension

Use these tools when the user wants to:

- inspect Outlook application status
- list default folders such as Inbox or Drafts
- inspect the active mail item
- list or search mail in the current or default folders
- create draft mail, reply, reply-all, or forward
- send an already prepared draft explicitly
- inspect or save attachments

Do not use this extension for non-Outlook formats or generic file editing outside Outlook workflows.

## Current seed categories

- `application`
- `folder`
- `mail`
- `attachment`

## Safety guidance

- Prefer inspection before mutation.
- Prefer draft-producing actions before `mail.send`.
- Inspect attachments before exporting them.
- Remember that some internal names still use inherited `PptMcp.*` migration naming.
