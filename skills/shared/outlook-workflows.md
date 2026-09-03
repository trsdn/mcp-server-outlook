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

## Not everything in a folder is a message

Every entry in a listing carries `itemType`:

- `mail` - an ordinary message
- `meetingRequest` - an invitation. Replying to it is **not** accepting it; a reply is just mail back
  to the organiser and leaves the invitation unanswered. Say so rather than implying you accepted.
- `meetingCancellation` - the meeting is off
- `meetingResponse` - somebody's answer to an invitation you sent
- `other` - something this surface does not model

`skippedItemCount` counts items that could not be summarised at all. If it is non-zero, the listing
is not the whole folder - do not describe it as such.

## Answering an invitation

`mail.respond-to-meeting` accepts, declines or tentatively accepts an invitation. Point it at the
invitation's `entryId` (its `itemType` is `meetingRequest`) with `response` set to `accept`,
`decline` or `tentative`.

Answering and notifying are separate steps, exactly as they are when creating a meeting:

- The response is always written to the user's own calendar.
- The organiser is told **only** when `sendResponse: true`. Add `responseText` to include a note.

So report which one you did. "I've accepted it" and "I've told them you're coming" are different
claims. Defaulting to not sending is deliberate: mail to a real organiser cannot be taken back.

The action refuses, by name, anything that is not an invitation - ordinary mail, a
`meetingCancellation` (the meeting is already off) and a `meetingResponse` (somebody else's answer to
a meeting the user organised). Do not work around a refusal by replying to the item instead; that
leaves the invitation unanswered.

## Inviting people is a separate step from creating the entry

`calendar.create-appointment` with `requiredAttendees` or `optionalAttendees` (semicolon-separated)
creates a **meeting** rather than a private appointment. It is saved to the user's own calendar and
**nobody is told**. Only `sendInvitation: true` mails the attendees.

So: say which one you did. "I've put it in your calendar" and "I've invited them" are different
claims, and only the second one requires `sendInvitation`.

If Outlook cannot resolve an attendee, the meeting is **not** created and `unresolvedAttendees` names
them. That is deliberate - an unresolved attendee never receives the invitation, so creating the
meeting anyway would look like success while leaving the person uninvited. Ask the user for a full
SMTP address rather than retrying the same name.

`calendar.read` reports `isMeeting` and an `attendees` list with each person's `responseStatus`
(`none`, `organizer`, `tentative`, `accepted`, `declined`, `notResponded`). `none` means the item is
not a meeting response yet - it does not mean they declined.

## Check availability before proposing a time

`calendar.get-free-busy` takes the same semicolon-separated attendee list and answers with each
person's `busyPeriods` - merged stretches of non-free time. Free time is everything they do not
cover. `availability` is Outlook's raw slot string behind that, one character per
`intervalMinutes`.

Two things will otherwise catch you out:

- **`end` is what was actually answered, not what you asked for.** Outlook decides how far ahead it
  publishes. When `end` is earlier than your window, `message` says so, and the time beyond it is
  unknown - not free.
- **An unresolvable attendee fails the call** rather than coming back free. Outlook reports an
  all-free calendar for somebody it never looked up, so a "free" answer for an unresolved name would
  be an invention.


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
