# OutlookMcp - Complete Feature Reference

**7 tools with 57 operations for Outlook automation**

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
| `move` | Move a message to another folder |
| `delete` | Delete a message |
| `set-read-state` | Mark a message read or unread |
| `set-flag` | Set or clear a follow-up flag on a message |
| `set-categories` | Set the categories on a message |
| `list-categories` | List the categories defined in the master category list |
| `list-rules` | List the Outlook rules defined on the store |
| `list-reminders` | List pending reminders |
| `set-subject` | Set the subject of a draft |
| `set-body` | Set the body of a draft |
| `set-recipients` | Set the recipients of a draft |
| `export` | Save a message to disk as `.msg`, `.txt`, `.html`, `.mht` or `.rtf` |

**`send` is the one irreversible operation in this surface.** It refuses to run without explicit
confirmation, and repeated calls carrying the same operation ID will not send twice. See #29.

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
| `create-appointment` | Create an appointment |
| `update-appointment` | Update an existing appointment |
| `delete-appointment` | Delete an appointment |
| `get-free-busy` | Read the free/busy availability of a recipient |
| `export` | Save an appointment to disk, including as iCalendar (`.ics`) |

**`.ics` is the calendar half of item export.** A mail item asked for iCalendar is refused, because
Outlook answers it with "Value does not fall within the expected range" - a message that reads like
an argument bug rather than "mail is not a calendar entry". The same absolute-path, overwrite and
format/extension rules as `mail export` apply.

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
| `delete` | Delete a folder |
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
| `delete` | Delete a contact |

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
| `delete` | Delete a task |

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

## Attachment Operations (4 operations)

| Action | Description |
|---|---|
| `list` | List the attachments on an item |
| `save` | Save an attachment to disk |
| `add` | Add an attachment to a draft |
| `remove` | Remove an attachment from a draft |

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
