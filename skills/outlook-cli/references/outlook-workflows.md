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
after the structured filters. By default it is exhaustive rather than indexed, so it will not miss a
term buried deep in a long message - but reaching the body means opening every candidate item, which
is slow, and the scan stops at a safety limit. Pair `query` with structured filters whenever you can:
they decide how many items have to be opened.

### Two search engines, and why the response says which one answered

`mail.search` takes `searchMode`:

- `clientScan` (**default**) - each candidate is opened and checked. **Substring** matching, so `foo`
  matches inside `foobar`. Exact, but bounded: in a folder larger than the scan limit a genuine match
  further back is never reached.
- `fullText` - Outlook's content index answers the query. **Whole-word** matching, so `foo` matches
  "a foo arrived" but *not* `foobar`. Nothing is opened client-side and there is no scan horizon, so
  it finds matches arbitrarily far back.

These are different questions, not fast and slow versions of the same one. Use `fullText` for a large
folder or a term you expect to be a real word; use the default when you need substring matching or a
short, exact answer.

Every search response reports `searchEngine` as `clientScan` or `contentIndex`. **Read it before you
conclude anything from an empty result** - the two engines disagree about what "no matches" means. If
you asked for `fullText` and the store could not serve it, `searchEngine` comes back `clientScan` and
`message` says why, rather than the tool pretending the index answered.


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

## Recurring series

Outlook stores a recurring series as a single master item dated at its **first** occurrence. So a
calendar listing that has not expanded the series shows a weekly stand-up on the week it started and
on no other week. Answering "are you free Tuesday?" from such a listing produces a confident yes for
a slot that is booked.

`calendar.list` expands a series into its individual occurrences **only when both `start` and
`endTime` are given** - a series with no end date has infinitely many occurrences, so an open-ended
range cannot be expanded. The result says which happened in `recurringExpanded`.

**Never conclude somebody is free from a listing whose `recurringExpanded` is false.** Re-list with
both bounds instead.

Each listed item carries `recurrenceState`:

- `notRecurring` - a one-off appointment
- `master` - the series itself, not a particular date
- `occurrence` - one instance of a series
- `exception` - an instance that was moved, shortened or otherwise changed

An occurrence carries the **master's** entry id, so editing or deleting it by entry id affects the
whole series, not just that date. To change or cancel one instance, name `occurrenceDate` - see below.

To create a series, pass `recurrenceType` (`daily`, `weekly`, `monthly` or `yearly`) to
`calendar.create-appointment`, with:

- `recurrenceInterval` - every N days/weeks/months/years, default 1
- `recurrenceDaysOfWeek` - semicolon-separated day names, weekly patterns only. Omitted, a weekly
  series repeats on the start day.
- `recurrenceCount` **or** `recurrenceEndDate` to bound it - not both, since Outlook keeps only one.
  Neither means it never ends.

`calendar.read` reports the stored pattern under `recurrence`, including `exceptionCount`. A non-zero
`exceptionCount` means the pattern alone does not describe the series - some occurrences differ - so
do not describe the schedule from the pattern without saying so.

### Changing or cancelling one instance

`calendar.update-appointment` and `calendar.delete-appointment` take an optional `occurrenceDate`.

- **Omitted** - the whole series is changed or cancelled. Every occurrence moves; every occurrence
  disappears.
- **Given** - only the instance starting on that date is touched. The rest of the series is left
  alone, and the changed instance becomes an `exception`.

The response reports `scope` as `series` or `occurrence`, so you can confirm what was actually
touched rather than inferring it from an entry id that is the same either way.

`occurrenceDate` may be a bare date (`2026-03-12`); the time of day is then taken from the series, so
you do not need to know that the stand-up starts at 09:17. Give a full date and time only when you
mean a specific instance that may already have been moved.

Naming `occurrenceDate` on an appointment that is not recurring is **refused**, not ignored - a
caller who asked to touch one instance of a series should not be told it succeeded against a one-off
item. A date the series does not fall on is likewise refused; check it against a listing first, and
remember that an instance somebody already cancelled no longer exists.

Cancelling a single occurrence of a **meeting** does not notify the attendees - nothing is sent by
this tool. Say so rather than implying the others have been told.

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

## More than one mailbox

A profile usually holds several stores: the main mailbox, an online archive, sometimes a second
account or an imported data file. **Every store has its own Inbox, Sent Items and Calendar.** An
unqualified request only ever reaches the *default delivery store*, so `folder.list-default` on its
own tells you about one mailbox and says nothing about the others existing.

```
1. folder.list-stores                                   → what mailboxes exist
2. folder.list-default(storeId: "<id from step 1>")     → that mailbox's well-known folders
3. mail.list(folder: "<folderPath from step 2>")         → read it
```

Address a store by `storeId`, never by `displayName` - two stores can share a name. A `storeId` that
does not resolve is refused rather than falling back to the default mailbox, because reading the
wrong account and reporting success is worse than failing.

Once you have a folder path from step 2 you can pass it straight to `mail.list`, `mail.search` and
`folder.list-children`: folder paths are absolute and already carry the store, so they reach any
store without further qualification.

Folder results name the store they came from (`storeId`, `storeName`). **If a user asks about mail
you cannot find, check whether you are looking at the right store before concluding it does not
exist** - an archived message is in a different store, not missing.

`list-stores` also reports `isDefaultStore`, `isDataFileStore`, `exchangeStoreType` and, where an
account delivers to the store, `accountSmtpAddress`. A store with no account is normally an archive
or an imported data file.

**A store does not necessarily have all the well-known folders.** An online archive typically has no
Inbox, Drafts or Calendar at all. `list-default` reports those roles as `available: false` with a
`note`; use `folder.list-children` on the store's `rootFolderPath` to see what it really contains.

Not covered: another person's mailbox by delegate access. Shared and delegate mailboxes that are not
already open in the profile are still unreachable.

## What this surface does not do

There is no window, visibility, or "watch me work" control. Outlook stays as the user left it.
Contacts, tasks, and rules are not exposed. If a user asks for one of these, say so plainly rather
than approximating it with mail operations.
