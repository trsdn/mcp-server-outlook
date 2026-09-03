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

## Narrow a search before you scan

`mail.list` and `mail.search` take structured filters that Outlook evaluates itself, before any item
reaches this server: `unreadOnly`, `fromAddress`, `subjectContains`, `receivedAfter`,
`receivedBefore` and `hasAttachment`. Prefer them over asking for a large list and filtering it
yourself. A structured filter reads only the matching items, so it finds mail that a plain listing
would never reach in a busy folder.

```
mail.list(folder: <inbox>, fromAddress: "anna@contoso.com", receivedAfter: "2024-03-01")
mail.list(folder: <inbox>, subjectContains: "invoice", hasAttachment: true)
```

Dates are ISO 8601 (`2024-03-07` or `2024-03-07T14:30`). A bare date means local midnight. All the
filters combine with AND.

`query` is different: it is a free-text match over subject, sender and the **full** body, applied
after the structured filters. It is exhaustive rather than indexed, so it will not miss a term buried
deep in a long message - but reaching the body means opening every candidate item, which is slow.
`Restrict` cannot filter on body text at all, so pair `query` with structured filters whenever you
can rather than relying on it alone in a large folder: the structured filters decide how many items
have to be opened.


## Read past the first page

A single `mail.list` or `mail.search` call returns at most one page. When the response has
`hasMore: true` there is more to see, and the only correct way to reach it is to pass `nextCursor`
straight back:

```
1. mail.list(folder: <inbox>, subjectContains: "invoice")   → hasMore: true, nextCursor: "..."
2. mail.list(folder: <inbox>, subjectContains: "invoice", cursor: <nextCursor>)
3. repeat until hasMore is false
```

Two rules make this safe:

- **Never conclude "there is no such mail" while `hasMore` is true.** A short or empty page means
  this call stopped early, not that the folder holds nothing further. Keep paging, or tell the user
  the search was incomplete.
- **A cursor only continues the query that produced it.** Keep `folder`, `query` and every filter
  identical across the walk; changing any of them is rejected, and you must restart without a
  cursor. `maxCount` is the exception - you may change page size mid-walk.

Do not try to build your own paging out of `receivedBefore`. Results are ordered by `receivedTime`
descending (`sortedBy` and `sortDirection` say so explicitly), and the cursor already handles
messages that share a received time; a hand-rolled date window silently drops them.


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

You cannot reply to or forward an unsent draft - there is nobody to reply to. To change a draft, use
`mail.set-subject`, `mail.set-body` or `mail.set-recipients` on the draft itself.

## Read a whole thread before answering

```
1. mail.search(query: ...) or mail.read-active     → find one message in the thread
2. mail.get-conversation(entryId: ...)             → the whole thread, oldest first
3. mail.read(entryId: ...)                         → full body of any item that matters
```

Do this before replying to anything that looks like part of an exchange. Answering from a single
message means answering without knowing what was already said - and the last message in a thread is
frequently the least informative one.

The thread spans folders: replies live in Sent Items, the original in the Inbox. Each item reports
its `folderPath`. `conversationId` also comes back on `mail.read`, `mail.list` and `mail.search`, so
you do not need an extra read to reach a thread.

If a store has conversation view disabled the call fails with `conversationSupported: false`. That
means *unknown*, not *no replies* - say so rather than implying the message stands alone.

A message you have only just created may not be listed in its thread yet: Outlook's conversation
index catches up a moment later. If you reply and immediately read the thread back to confirm, a
missing reply means "not indexed yet", not "the reply was lost". Check the draft with `mail.read`
before telling a user something went wrong.

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

`list-items` returns the newest items first and caps at `maxCount`. When it reports `truncated: true`
you are looking at the newest slice of a larger folder, not the whole of it - do not conclude an item
is absent from a truncated listing. Check `sortedBy`: if it is null the folder had no orderable
timestamp and the order is arbitrary, so a truncated result there tells you nothing about what is
missing.

## What this surface does not do

There is no window, visibility, or "watch me work" control. Outlook stays as the user left it.
Contacts, tasks, and rules are not exposed. If a user asks for one of these, say so plainly rather
than approximating it with mail operations.
