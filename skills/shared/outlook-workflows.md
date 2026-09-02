# Common Outlook Workflows

Worked sequences for the tasks users ask for most. All calls are shown in MCP form; the CLI takes the
same action and parameter names (`outlookcli mail list --folder Inbox`).

## Triage the inbox

```
1. application.get-status                          → confirm classic Outlook is available
2. folder.list-default                             → get the Inbox folder
3. mail.list(folder: <inbox>, unreadOnly: true)    → the working set
4. mail.read(entryId: ...)                         → only for items that need the body
5. mail.set-categories / mail.set-read-state / mail.move
```

Summarise the set for the user before acting on it in bulk. Confirm before anything destructive.

## Reply to the message the user is looking at

```
1. mail.read-active                                → the open or selected item
2. mail.reply(entryId: ...)                        → creates a draft, does not send
3. mail.set-body(entryId: <draft>, body: ...)
4. show the user the draft
5. mail.send(entryId: <draft>, confirm: true)      → only after they say yes
```

Use `reply-all` only when the user asked for it. Defaulting to reply-all is a common and expensive
mistake.

## Compose a new message

```
1. mail.create-draft(...)
2. mail.set-recipients / mail.set-subject / mail.set-body
3. attachment.add(entryId: <draft>, path: ...)     → if needed
4. read the draft back to the user
5. mail.send(entryId: <draft>, confirm: true)
```

## Save attachments

```
1. mail.search(...) or mail.list(...)              → find the message
2. attachment.list(entryId: ...)                   → see what is on it
3. attachment.save(entryId: ..., index: ..., path: ...)
```

Ask where to save if the user did not say. That is a genuine preference, so Rule 2 does not apply.

## Look at the calendar

```
1. calendar.list(start: ..., end: ...)             → appointments in a range
2. calendar.read(entryId: ...)                     → details for one
3. calendar.create-appointment / update-appointment / delete-appointment
```

Confirm before `delete-appointment`, and before an `update-appointment` that changes the time of a
meeting with other attendees.

## Walk a folder tree

```
1. folder.list-default                             → the well-known folders
2. folder.list-children(folder: ...)               → descend
3. folder.resolve-path(path: "Inbox/Projects/Q1")  → if the user gave you a path
4. folder.list-items(folder: ...)
```

`resolve-path` is cheaper than walking the tree when the user already named the folder.

## What this surface does not do

There is no window, visibility, or "watch me work" control. Outlook stays as the user left it.
Contacts, tasks, and rules are not exposed. If a user asks for one of these, say so plainly rather
than approximating it with mail operations.
