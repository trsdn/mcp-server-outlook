# OutlookMcp - Complete Feature Reference

**10 tools with 69 operations for Outlook automation**

This document is derived from the generated `ServiceRegistry` action lists, which are the single
source of truth for the tool surface. Both entry points expose exactly these operations:

- **MCP Server** - conversational tool surface with rich tool schemas
- **CLI** (`outlookcli`) - compact scripting and coding-agent surface

Every action below is available identically through both. If you find one that is not, that is a
bug (see Rule 24, post-change sync).

---

## Mail Operations (23 operations)

| Action | Description |
|---|---|
| `read-active` | Read the mail item currently open or selected in Outlook |
| `read` | Read a mail item by entry ID |
| `list` | List mail items in a folder |
| `search` | Search mail items |
| `get-conversation` | Read a whole conversation thread. Non-mail items on the thread are returned in `otherItems` rather than dropped |
| `respond-to-meeting` | Accept, tentatively accept or decline a meeting invitation |
| `create-draft` | Create a draft message |
| `reply` | Reply to the sender of a message |
| `reply-all` | Reply to all recipients of a message |
| `forward` | Forward a message |
| `send` | Send a draft. Requires explicit confirmation and is idempotent per operation ID |
| `move` | Move a message to another folder. Recoverable, so ungated |
| `delete` | Delete a message. Soft delete to Deleted Items, so ungated - unless the message is already there, which is permanent and requires confirmation |
| `set-read-state` | Mark a message read or unread |
| `set-flag` | Set or clear a follow-up flag on a message |
| `set-categories` | Set the categories on a message |
| `list-categories` | List the categories defined in the master category list |
| `list-rules` | List the Outlook rules defined on the store. An alias for `rule list`, kept here because "why is nothing arriving from this sender?" is a mail question whose answer is a rule |
| `list-reminders` | List pending reminders |
| `set-subject` | Set the subject of a draft |
| `set-body` | Set the body of a draft |
| `set-recipients` | Set the recipients of a draft |
| `export` | Save a message to disk as `.msg`, `.txt`, `.html`, `.mht` or `.rtf` |

**`send` is the one operation here whose effect leaves the mailbox entirely.** It refuses to run
without explicit confirmation, and repeated calls carrying the same operation ID will not send
twice. See #29.

**Confirmation gates are drawn at recoverability, not at how alarming the verb is** (#9). An action
takes `confirm` only where Outlook offers no way back:

| Gated - refused without `confirm=true` | Not gated - recoverable |
|---|---|
| `mail.send` | `mail.delete` (moves to Deleted Items) |
| `folder.delete` (takes every message and subfolder; not a recycle-bin operation in every store) | `mail.move` (move it back) |
| `attachment.remove` (an attachment has no Deleted Items of its own) | `contact.delete`, `task.delete` (move to Deleted Items) |
| `calendar.delete-appointment` **with `occurrenceDate`** (writes a deletion exception into the recurrence pattern) | `calendar.delete-appointment` for the whole appointment or series |
| any item delete whose target **is already in Deleted Items** - there is no second recycle bin | |

Gating a recoverable action would train a caller to pass `confirm=true` reflexively, which is how
the gate on the irreversible one stops being read. The ungated actions are a decision, not an
oversight, and they still require the caller to report what was deleted and where it went.

**`export` writes Unicode `.msg`, never the ANSI variant.** Outlook's `olMSG` silently replaces any
character outside the machine's code page with `?` and reports success, so `msg` always means
`olMSGUnicode` here. `filePath` must be absolute: Outlook accepts a relative path and resolves it
against its own working directory rather than the caller's. An existing file is never replaced
unless `overwrite` is set, and a format that contradicts the file extension is refused rather than
producing, say, a binary `.msg` under a `.txt` name.

---

## Calendar Operations (7 operations)

| Action | Description |
|---|---|
| `list` | List appointments in a date range |
| `read` | Read an appointment by entry ID |
| `create-appointment` | Create an appointment, or a meeting with attendees and room resources |
| `update-appointment` | Update an existing appointment |
| `delete-appointment` | Delete an appointment or cancel one occurrence. Deleting the appointment or series is ungated (it goes to Deleted Items); cancelling a single `occurrenceDate` requires confirmation |
| `get-free-busy` | Read the free/busy availability of a recipient |
| `export` | Save an appointment to disk, including as iCalendar (`.ics`) |

**`.ics` is the calendar half of item export.** A mail item asked for iCalendar is refused, because
Outlook answers it with "Value does not fall within the expected range" - a message that reads like
an argument bug rather than "mail is not a calendar entry". The same absolute-path, overwrite and
format/extension rules as `mail export` apply.

**Rooms are booked with `resourceAttendees`, not `location`.** `location` is a free-text label that
reserves nothing, and a room named in `requiredAttendees` is invited like a person rather than
booked. Note that resolution is a weak existence check: any SMTP-shaped address resolves as a one-off
whether or not the mailbox exists, so a successful create is not proof that the room is real.

---

## Folder Operations (10 operations)

| Action | Description |
|---|---|
| `list-default` | List the default Outlook folders |
| `list-stores` | List the stores (accounts and PST files) attached to the profile |
| `open-shared` | Open a folder in another user's mailbox |
| `create` | Create a folder |
| `rename` | Rename a folder |
| `move` | Move a folder under a different parent |
| `delete` | Delete a folder together with everything filed in it. Requires confirmation |
| `list-children` | List the child folders of a folder |
| `resolve-path` | Resolve a folder path to a folder |
| `list-items` | List the items in a folder |

---

## Contact Operations (5 operations)

| Action | Description |
|---|---|
| `list` | List the contacts in the default Contacts folder or an explicit folder path. Distribution lists are returned separately in `distributionLists` rather than dropped |
| `read` | Read a contact by entry ID, or the contact currently open or selected in Outlook |
| `create` | Create a contact |
| `update` | Update named fields on an existing contact. Fields that are not passed are left alone |
| `delete` | Delete a contact. Ungated soft delete, unless the contact is already in Deleted Items |

A Contacts folder holds distribution lists as well as people. `contacts`, `distributionLists` and
`skippedItemCount` together always account for every item scanned, so nothing can be silently
dropped. Names are not unique and some contacts have no name at all, so `entryId` is the only
reliable handle.

---

## Task Operations (5 operations)

| Action | Description |
|---|---|
| `list` | List the tasks in the default Tasks folder or an explicit folder path. Completed tasks are omitted unless `includeCompleted` is true |
| `read` | Read a task by entry ID, or the task currently open or selected in Outlook |
| `create` | Create a task |
| `update` | Update named fields on an existing task. Fields that are not passed are left alone |
| `delete` | Delete a task. Ungated soft delete, unless the task is already in Deleted Items |

Two things about real task folders shape this surface, and both were measured rather than assumed.

**Outlook does not use null for "no date".** An unset `dueDate`, `startDate` or `dateCompleted`
reads as 1 January 4501 through COM - on the mailbox this was built against, 260 of 274 due dates.
That sentinel is never returned: a missing `dueDate` means there is no due date, not a due date in
the 46th century.

**Nearly every task is already finished** - 271 of those 274 - so `list` omits completed tasks by
default. `completedItemCount` reports how many were filtered out, so "no open tasks" can never be
confused with "this listing dropped rows".

Set `status` to `complete` to mark a task done; Outlook then sets `percentComplete` to 100 and
stamps `dateCompleted` itself. `status` is one of `not-started`, `in-progress`, `complete`,
`waiting` or `deferred`.

---

## Rule Operations (5 operations)

| Action | Description |
|---|---|
| `list` | List a store's rules. `includeDetail` adds each rule's conditions, actions, subject terms, sender addresses and move-to destination |
| `create` | Create a rule. Requires at least one condition and at least one action |
| `update` | Change an existing rule's clauses. A clause that is not passed is left alone; an empty string clears one |
| `set-enabled` | Switch a rule on or off |
| `delete` | Remove a rule |

**Rule writes are the highest-risk operation in this surface, above `mail delete`.** A message
deleted in error sits in Deleted Items; a rule created in error silently moves or destroys mail that
has not arrived yet, keeps doing it, and is typically noticed days later.

**Rules are per-store.** Every action defaults to the profile's default delivery store and takes a
`storeId` from `folder list-stores` to reach another mailbox. An unknown `storeId` is refused rather
than falling back to the default store, because rewriting the wrong mailbox's rules under
`success: true` is the worst outcome available here.

**Rules are addressed by name, and Outlook permits duplicates.** `create` therefore refuses a name
already in use, and `update`, `set-enabled` and `delete` refuse a name matching no rule or more than
one, rather than picking the first.

**A rule with no conditions matches every message that arrives, and one with no actions does
nothing.** Outlook accepts both; this refuses both, on create and on update.

**Outlook inserts a new rule at the top of the evaluation order, not the bottom** - verified against
a live mailbox, where a newly created rule came back with `executionOrder` 1. A new rule therefore
runs before every rule the mailbox already had.

**There is no mark-as-read action, and this is not an omission.** Outlook's rule object model has no
such action - `RuleActions` exposes `AssignToCategory`, `CC`, `ClearCategories`, `CopyToFolder`,
`Delete`, `DeletePermanently`, `DesktopAlert`, `Forward`, `ForwardAsAttachment`, `MarkAsTask`,
`MoveToFolder`, `NewItemAlert`, `NotifyDelivery`, `NotifyRead`, `PlaySound`, `Redirect` and `Stop`,
and nothing else. Only the Rules and Alerts wizard inside Outlook can create one.

**`deleteMessage` does not read back as a delete.** Outlook has no delete action either: it rewrites
"delete it" into a move to Deleted Items plus stop-processing, so `list` afterwards reports
`moveToFolder` with a Deleted Items destination. For the same reason `deleteMessage` and
`moveToFolder` cannot both be set - a rule has one move destination.

**Deliberately out of scope**, and each for a reason rather than for want of time:

| Not exposed | Why |
|---|---|
| `Forward`, `Redirect`, `CC` | They send mail on the user's behalf, unattended, indefinitely. Configuring that in a single tool call is not a capability an agent should have. They also need `Recipients.ResolveAll`, which is Object Model Guard-protected |
| `DeletePermanently` | Unrecoverable, with no Deleted Items to retrieve from |
| The `From` condition | Holds address-book entries and needs `Recipients.ResolveAll`, which raises the Object Model Guard prompt that cannot be answered programmatically. `SenderAddress` matches the SMTP address directly and is what "mail from this person" almost always means. Existing `From` rules are still read back correctly by `list` |
| Multi-term conditions | A rule matching several subject terms is read back correctly but written with one term, to keep the argument shape unambiguous |
| Exceptions, `Account`, `Importance`, `MessageHeader`, `FormName`, RSS conditions, `PlaySound`, `DesktopAlert`, `MarkAsTask` | Enumerable through `list`, but writing them is a long tail with no agent use case that justifies the surface |
| Send rules (`olRuleSend`) | A different mental model - they fire on messages the user sends. Enumerated by `list` as `ruleType: send`, not creatable |
| Reordering (`ExecutionOrder`) | Changing evaluation order rewrites the meaning of every rule that stops processing, with no way to preview the effect |

**Writes are collection-wide, and only the save persists anything.** Outlook commits a store's whole
rule collection at once, so every write here rewrites all of the mailbox's rules; the response
reports `ruleCount` so a caller can check the total is what they expected. That save is the step
that fails - on the Exchange rules quota, on the user having the Rules and Alerts wizard open, or on
some unrelated rule in the mailbox being malformed - and when it fails nothing was written at all,
not even partially.

---

## Attachment Operations (4 operations)

| Action | Description |
|---|---|
| `list` | List the attachments on an item |
| `save` | Save an attachment to disk |
| `add` | Add an attachment to a draft |
| `remove` | Remove an attachment from a draft. Requires confirmation: an attachment has no Deleted Items to be recovered from |

---

## Application Operations (3 operations)

| Action | Description |
|---|---|
| `get-status` | Report Outlook availability, including whether the installed client is classic Outlook or the new Outlook (#35) |
| `get-active-explorer` | Report which folder the user is looking at and what is selected there |
| `get-active-inspector` | Report which item the user currently has open |

`get-status` is the right first call in any workflow. The new Outlook does not expose a COM object
model, so every other operation in this document requires classic Outlook.

`get-active-explorer` and `get-active-inspector` both answer "nothing is open" as a success, not an
error. An item the user is still composing has not been saved and therefore has no `entryId`, so it
cannot be addressed by any other action until it is; `isSaved` says which case you are in.

---

## Address Book Operations (3 operations)

| Action | Description |
|---|---|
| `resolve` | Check one or more addressees against the address book and report their real SMTP addresses |
| `list-address-lists` | List the address books attached to the profile: the Global Address List, Contacts, LDAP directories |
| `list-entries` | Browse the entries in one address book |

**`resolve` is the check to run before sending.** It answers per addressee, and `allResolved` is
the single flag to test; `unresolvedNames` says which ones are wrong. A name Outlook cannot find is
a success with `resolved: false`, not an error - "no such person" and "Outlook could not be
reached" are different answers and must not collapse into one. An ambiguous name also comes back
unresolved: Outlook's object model offers no way to list the candidates, so pass the full SMTP
address to disambiguate.

**Semicolons separate addressees; commas do not.** `Smith, Jane` is one addressee. That is the usual
Global Address List display-name shape, so splitting on commas would take the commonest form of the
exact input this action exists to resolve and turn it into two fragments that resolve to nothing.
Outlook separates recipients with `;` for the same reason.

**`smtpAddress` is always a mailable address.** Outlook's own `AddressEntry.Address` returns an
X500 legacyExchangeDN - `/o=ExchangeLabs/ou=.../cn=Recipients/cn=...` - for an Exchange entry. It
is a string, it serialises cleanly, and mail sent to it goes nowhere. That value is reported
separately as `rawAddress` and is never passed off as an email address. `smtpAddressSource` says
which route produced the answer: the Exchange directory, a distribution list, a local contact, a
`PR_SMTP_ADDRESS` read, or a one-off SMTP string that was never checked against anything.

**`list-entries` scans; it does not search.** The Outlook object model has no `Restrict` or `Find`
on an address book, so `startsWith` is applied while scanning and the scan stops at `scanLimit`. A
corporate Global Address List is far larger than that, so check `scanLimitReached`: when it is
true, an empty result is not evidence that nobody matches, and `resolve` is the right call for
someone you can already name.

The scan starts at the beginning of the book and does not jump to a prefix. Measured on a real
corporate GAL: scanning 3000 entries for names starting with `S` matched **none of them**, because
the first 3000 entries begin with punctuation and digits. The prefix filter is genuinely useful
against a Contacts folder, which fits inside the budget; against a GAL, `resolve` is the only
realistic way to find a person.

**Every action here is Object Model Guard territory.** Recipients and address entries are exactly
the members Outlook protects against out-of-process callers, so any of these calls can be refused
by a modal security prompt that no program can answer. A refusal fails the call with an
explanation; a property refused while the rest of the call succeeded is named in `accessDenied`, so
a missing value is never confused with a value the directory does not hold.

That distinction is load-bearing here rather than merely tidy. This surface exists to validate an
addressee *before* sending, and "Outlook has no such person" and "Outlook refused to tell me" call
for opposite actions - correct the address, or treat the answer as unknown and do not call the send
validated. `accessDenied` covers the protected members that matter: `Recipient.AddressEntry`,
`AddressEntry.Address`, `GetExchangeUser`, `GetExchangeDistributionList`, `GetContact`,
`ExchangeUser.PrimarySmtpAddress` and `PropertyAccessor`.

---

## Message Property Operations (4 operations)

| Action | Description |
|---|---|
| `get-headers` | Read the internet message headers of a received message, parsed into names and values |
| `get-known` | Read a curated set of MAPI properties that are commonly useful and awkward to get right by hand |
| `get-property` | Read any MAPI property by its DASL name |
| `list-user-properties` | List the custom user properties on an item |

Read-only. There is no way to write a property through this tool.

**A draft has no transport headers.** Nothing composed locally ever traversed an SMTP transport, so
it carries none, and the call succeeds with `headersPresent: false`. The same is often true of an
item delivered entirely inside one organisation. That is an answer, not a failure. `headersPresent`
is also false when Outlook refused the read, so check `status` before concluding a message has no
headers: "there are none" and "Outlook would not say" are different claims.

**Headers are unfolded.** An RFC 5322 continuation line begins with whitespace and continues the
header above it, so a line-by-line split invents nameless entries and truncates exactly the headers
worth reading - `Received` and `Authentication-Results` are almost always folded. Duplicates are
preserved in transport order, because a message carries one `Received` header per relay hop and
their order is the delivery path in reverse. A header block runs to tens of kilobytes, so
`headerName` returns one header rather than all of them and `includeRaw` is off by default.

**Absence has two shapes, and both mean "no usable value".** Outlook raises `MAPI_E_NOT_FOUND` when
an item does not carry a property, which is reported as `not-present`. But an Exchange store returns
an **empty string** for some tags rather than reporting them missing - `PR_TRANSPORT_MESSAGE_HEADERS`
and `PR_INTERNET_MESSAGE_ID` on a draft both do - and reporting that as a found value would answer
"yes, this message has an Internet message id" while handing back nothing. That case is `empty`.
`found` is false for both, so one check answers "is there a value here"; the status says which.

**A refusal is not an absence.** `blocked` means the value exists and Outlook withheld it.
`unsupported-or-blocked` is Outlook's `MAPI_E_NOT_SUPPORTED`, which is genuinely ambiguous: it is
returned both for a property type the accessor cannot handle at all (`PT_OBJECT`) and for a security
refusal, and the HRESULT alone cannot tell them apart. It is reported as ambiguous rather than
asserted to be one or the other.

**`get-property` reads any MAPI property, not a curated list.** This is a deliberate choice. It is
read-only and it cannot reach an item the caller could not already open in full with `mail.read`, so
it grants no access the rest of this surface does not already grant - but it does expose properties
this surface deliberately does not project, and that is stated here rather than left to be
discovered. A fixed allow-list over a property space with thousands of members would be permanently
incomplete, and the curation would be guesswork. `get-headers` and `get-known` exist so that the
common questions do not need it.

**Binary properties are not stringified.** A `PT_BINARY` value arrives as a byte array; anything
that calls `ToString()` on it emits the literal `System.Byte[]` and reports success. Binary values
come back as base64 and as the hex form Outlook itself uses, which is the form an entry id has to be
in to be handed back to Outlook.

---

## Requirements

- Windows
- **Classic** Outlook, installed and configured with a profile

Outlook COM is invoked through a dedicated STA dispatcher, so concurrent requests are serialized
onto a single apartment-threaded worker. See `docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md`.

---

## What this document no longer covers

Earlier revisions of this file catalogued 33 tools and 204 operations for presentation automation:
slides, shapes, text, charts, animations, transitions, SmartArt, VBA and more. That surface was
inherited when this repository was renamed, and it was deleted in #26.
None of those operations exist any more, in either the MCP server or the CLI.

The inherited presentation-session COM plumbing (`ComInterop/Session/*`) has also been deleted.
Outlook has no document to open or save, so there is no session or batch concept anywhere in the
product. See ADR-002 for the reasoning.
