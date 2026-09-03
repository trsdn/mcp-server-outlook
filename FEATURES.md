# OutlookMcp - Complete Feature Reference

**6 tools with 48 operations for Outlook automation**

This document is derived from the generated `ServiceRegistry` action lists, which are the single
source of truth for the tool surface. Both entry points expose exactly these operations:

- **MCP Server** - conversational tool surface with rich tool schemas
- **CLI** (`outlookcli`) - compact scripting and coding-agent surface

Every action below is available identically through both. If you find one that is not, that is a
bug (see Rule 24, post-change sync).

---

## Mail Operations (22 operations)

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

**`send` is the one irreversible operation in this surface.** It refuses to run without explicit
confirmation, and repeated calls carrying the same operation ID will not send twice. See #29.

---

## Calendar Operations (6 operations)

| Action | Description |
|---|---|
| `list` | List appointments in a date range |
| `read` | Read an appointment by entry ID |
| `create-appointment` | Create an appointment |
| `update-appointment` | Update an existing appointment |
| `delete-appointment` | Delete an appointment |
| `get-free-busy` | Read the free/busy availability of a recipient |

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

## Attachment Operations (4 operations)

| Action | Description |
|---|---|
| `list` | List the attachments on an item |
| `save` | Save an attachment to disk |
| `add` | Add an attachment to a draft |
| `remove` | Remove an attachment from a draft |

---

## Application Operations (1 operation)

| Action | Description |
|---|---|
| `get-status` | Report Outlook availability, including whether the installed client is classic Outlook or the new Outlook (#35) |

`get-status` is the right first call in any workflow. The new Outlook does not expose a COM object
model, so every other operation in this document requires classic Outlook.

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
