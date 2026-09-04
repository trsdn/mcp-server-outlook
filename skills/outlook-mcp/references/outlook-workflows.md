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

## Formatted mail, and when not to use it

`create-draft`, `reply`, `reply-all`, `forward` and `set-body` all take `bodyFormat`, which is
`plain` (the default) or `html`.

Pass `html` when the body really is markup you want rendered - a bulleted list, a link, bold text, a
table:

```
mail.create-draft(subject: "Status", body: "<ul><li>Design signed off</li><li>Build starts Monday</li></ul>", bodyFormat: "html")
```

Leave it as `plain` for ordinary prose, and in particular whenever the text came from the user rather
than from you. Plain text is escaped, not interpreted, so `profit < loss` arrives as written. Send
that same string as `html` and everything from the `<` onwards disappears into what the renderer
takes for an unclosed tag - the message still sends, and still reports success, and quietly says
something other than what the user asked you to say.

An unrecognised `bodyFormat` is refused rather than treated as plain, so a typo fails loudly instead
of putting raw tags in front of a human.

On `reply`, `reply-all` and `forward` the body goes *above* the quoted original, and the quoted
message keeps its own formatting whichever format you choose - you do not have to match it.

## Follow-up flags

Every listing and read reports `flagStatus`, always, as `none`, `flagged` or `complete`. It is never
omitted, so "this message is not flagged" is distinguishable from "this listing does not tell you
about flags" - the two would otherwise look identical and only one of them means there is nothing to
do. `flagRequest` (the label) and `flagDueDate` appear when they are set.

That means "what still needs following up?" is answerable from a single `mail.list` call. Do not read
each message in turn to find out, and do not list everything and filter client-side either - pass
`flaggedOnly: true` and Outlook does the work over the folder before anything is handed back.

```
1. mail.list(folder: "inbox", maxCount: 100, flaggedOnly: true)
2. mail.set-flag(entryId: ..., flagStatus: "complete")
```

`flaggedOnly` returns **outstanding** flags only. A completed flag is finished work, so it is
excluded - otherwise every item the user has already dealt with would come straight back onto their
to-do list. It combines with the other filters, so `flaggedOnly` plus `fromAddress` answers "what am
I still on the hook for from this person".

`complete` and `none` are not interchangeable. `complete` says the work was done; `none` says it was
never raised. Clearing a flag the user has finished with throws away the record that they finished
it, so prefer `complete` unless they actually want the flag removed.

Two limits Outlook imposes, not this server:

- A flag can only be **completed** on a message that has been sent or received. Completing a draft is
  refused with an explanation rather than a raw COM error.
- Flagging a draft works, but a draft is rarely what you want to flag in the first place.

An unrecognised `flagStatus` is refused rather than guessed at, and nothing is applied.

## Categories: discover before you write

Outlook does **not** validate what `mail.set-categories` writes. Assigning a name that is not in the
mailbox's list succeeds, returns `success: true`, and produces a category the user cannot filter or
colour by. So a typo does not fail - it quietly creates a dead label.

Read the list first and use a name from it:

```
1. mail.list-categories              → { name: "compliance", color: "darkPeach" }, ...
2. mail.set-categories(entryId: ..., categories: "compliance")
```

Colours come back as names (`yellow`, `darkTeal`, or `none` for a category created without one),
never as raw enum numbers, so you can repeat them back to the user as they appear in Outlook.

If the user names a category that is not in the list, say so and offer the closest matches rather
than writing it anyway. Creating categories is not exposed here - that is a mailbox-wide setting.

`set-categories` replaces the whole set on the message, so include the existing categories from
`mail.read` if the intent is to add one rather than to replace them all.

## Reminders are mostly overdue

`mail.list-reminders` returns what Outlook intends to remind the user about - appointments, tasks and
flagged mail together - earliest first.

The thing to know before reporting anything from it: on a long-lived mailbox **most reminders are
overdue**, often by years. On the mailbox this was built against, 416 of 605 had already passed, the
oldest by five years. So they are excluded by default, and the result tells you how many were held
back:

```
mail.list-reminders()   → { reminders: [...],           // upcoming, earliest first
                            totalCount: 605,
                            upcomingCount: 189,
                            overdueCount: 416 }
```

Report `overdueCount` when it is non-zero. A user shown fifty upcoming reminders and not told that
four hundred have already lapsed has been given a tidy and misleading picture. Pass
`upcomingOnly: false` to see them, and `maxCount` to widen the page - the counts always describe the
whole set, not the page.

`isOverdue` is derived from the due time. Do not look for a "pending" flag on the item: Outlook's
`IsVisible` means *the reminder dialog is on screen right now*, which is false for essentially every
reminder, and reading it as pending-ness would report a mailbox stacked with reminders as having none.

Read-only. Dismissing and snoozing are not exposed.

## Rules explain missing mail

Before telling a user that a folder is empty or that nothing arrived from someone, check whether a
rule already moved it. Rules run before this tool sees anything, so "no mail from Anna" and "a rule
files Anna's mail into Projects" look identical from a listing.

```
1. mail.list-rules(includeDetail: true)   → { name: ..., conditions: ["from"],
                                              fromAddresses: ["anna@..."],
                                              actions: ["moveToFolder", "stop"],
                                              moveToFolderPath: "\\...\Inbox\Projects" }
2. mail.list(folder: "Inbox\Projects")
```

`includeDetail` is **off by default** because it is roughly forty times the work - Outlook stores a
fixed slot for every condition and action it supports, so detail means walking about sixty slots per
rule. Use the plain listing when you only need names and whether they are on.

Things worth reading off the result:

- `enabled: false` means the rule explains nothing about where mail went.
- `executionOrder` matters, because a `stop` action means later rules never ran.
- `isLocalRule: true` means the rule only runs while Outlook is open, which is a common reason mail
  is filed late or not at all.
- `ruleType` is `receive` or `send`. A send rule does not explain missing incoming mail.

Rules are read-only here. Creating or changing one alters real mail flow for every future message,
so if the user wants that, tell them where it is in Outlook rather than implying this tool can do it.

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

### Someone else's mailbox


The stores above are the ones already in the profile. A shared or delegate mailbox that is *not* in
the profile is reached by address instead:

```
folder.open-shared(address: "team-inbox@contoso.com", role: "calendar")
  → folderPath, then mail.list / calendar.list against that path
```

`role` is one of `inbox`, `calendar`, `contacts`, `tasks`, `notes`, `journal` - Outlook supports no
others for a mailbox opened this way. The returned `folderPath` behaves like any other: pass it to
`mail.list`, `calendar.list` or `folder.list-children`.

**`address` is required and is never guessed.** Outlook resolves an unknown address to the current
user, so a defaulted or mistyped address would return *your own* inbox with `success: true` - a
convincing wrong answer, which is worse than an error. If the address cannot be opened the call
fails and names the mailbox.

A failure here means either no such mailbox or no access granted; **Outlook cannot distinguish
those**, so the error says both. If a user expects access they do not have, the fix is on the
Exchange side, not here.

## Changing the folder tree

`folder.create` takes a parent and a name; `rename`, `move` and `delete` take the folder itself.
Paths returned by any of them are ordinary folder paths and work with `mail.list`, `mail.move` and
the rest.

```
1. folder.create(parentFolder: "inbox", name: "Archive 2024")   → folderPath
2. mail.move(entryId: "...", destinationFolder: "<folderPath>")
```

**`delete` takes the folder's contents with it.** Everything filed in it, and every subfolder, goes
too. There is no undo beyond whatever Deleted Items happens to retain. Say what will be deleted
before deleting it.

**Default folders and store roots are refused** for `rename`, `move` and `delete` - Inbox, Sent
Items, Drafts, Deleted Items, Calendar, Contacts, Tasks, Notes, Junk, Outbox, in *every* store, not
just the default one. Outlook itself permits deleting the Inbox: no prompt, no error, and the mail
goes with it. If a user asks for that, tell them it is refused and why rather than looking for a way
round it.

A folder merely *named* "Inbox" that is not the default Inbox is an ordinary folder and can be
deleted normally - the check compares identity, not names.

Not covered: emptying a folder in place. Delete the items with `mail.delete`, or delete the folder.

## What this surface does not do

There is no window, visibility, or "watch me work" control. Outlook stays as the user left it.

## Task dates are 4501, not null

`task.list` and `task.read` are the Outlook Tasks folder. Two things about real task folders will
mislead you if you do not know them.

**Outlook stores 1 January 4501 for a date that was never set.** The tools strip that sentinel, so a
task with no `dueDate` simply has no `dueDate` field - it does not have a due date in the 46th
century. If you ever see a 4501 date reach a user, that is a bug worth reporting, not a real date.
Do not invent a due date for a task that has none: "no due date" is a normal and common state. On the
mailbox this was built against, 260 of 274 tasks had no due date at all.

**Most tasks in a real folder are already finished**, so `list` omits completed ones by default. That
is a filter, and it is reported: `completedItemCount` says how many were hidden and
`includedCompleted` echoes the flag back. If a user asks *what have I completed?* or *show me
everything*, pass `includeCompleted: true` - otherwise you will answer "you have three tasks" about a
folder holding 274.

`status` is `not-started`, `in-progress`, `complete`, `waiting` or `deferred`. To mark something
done, set `status` to `complete`; Outlook fills in `percentComplete` and `dateCompleted` itself, so
do not set them by hand as well. `subject` is not unique - `entryId` is the only reliable handle,
exactly as with contacts.

Not covered: task requests (assigning a task to someone else), and task recurrence.

## A Contacts folder is not all contacts

`contact.list` returns two lists, and a Contacts folder holds distribution lists as well as people.

`contacts` holds the people. `distributionLists` holds the groups, each with its `memberCount`.
*Who is in the team list?* is answered entirely by `distributionLists`, so a count built from
`contacts` alone will be confidently short. On a real address book of 83 items, 82 were people and
one was a distribution list of 13 members.

`contacts.length + distributionLists.length + skippedItemCount` always equals `scannedItemCount`.
If it does not, something was dropped and the result is not trustworthy - say so.

Two further traps:

- `totalItemCount` is the size of the folder; `scannedItemCount` is how many items were examined.
  When `truncated` is true they differ, and the listing is a first page, not the address book.
- **Some contacts have no name.** Real address books contain entries whose first name, last name,
  company and email are all blank; those come back as `(contact with no name)`. Names are not unique
  either. `entryId` is the only reliable handle - use it for `read`, `update` and `delete` rather
  than matching on a display name.

`update` writes only the fields you pass. Omitting a field leaves it alone; passing an empty string
clears it. To blank a phone number, pass `""`, not `null`.

**Outlook rewrites what you store.** A phone number written as `+1 555 0100` reads back as
`+1 (555) 0100`, because Outlook canonicalises it on save. Do not report a write as failed because
the value you read back differs from the value you sent, and do not compare contacts by string
equality on a phone number.

## A thread is not all mail

`get-conversation` returns two lists, and reading only the first one loses most of the answer.

`messages` holds the mail. `otherItems` holds the meeting invitation, the calendar appointment it
created, and the acceptances or declines - each with a named `itemType` and the folder it lives in.
On a real thread these were four of seven items. *When did we agree to meet, and did they accept?*
is answered entirely by `otherItems`, so a summary built from `messages` alone will confidently
omit it.

`skippedItemCount` means something narrower: entries the conversation still lists but the store
could not return - deleted mid-read, or in a store this profile cannot open. It is a data-loss
signal. If it is non-zero, say so rather than presenting the thread as complete.

## Do not guess which folder the user means

`application get-active-explorer` reports the folder the user is currently looking at, its full store
path, and how many items are selected there. When a request says "this folder", "here", or "the
messages I have selected", ask Outlook rather than guessing at a folder name - a guess that lands on
the wrong store fails in a way that looks like an empty mailbox.

`application get-active-inspector` reports the item the user has open. Two things about it are easy
to get wrong:

- An item the user is still composing has never been saved, so it has **no `entryId`**. The field is
  omitted and `isSaved` is `false`. Without an `entryId` no other action can address that item; if
  you need to act on it, say that the user must save or send it first. Do not invent an identifier
  and do not fall back to matching on the subject.
- The parent folder of an unsaved outgoing item reads as the **Outbox**, not Drafts. That is where it
  would go, not where it is. Do not report it to the user as already being in a folder.
