# Behavioral Rules for Outlook Operations

These rules apply to every Outlook operation, through both the MCP server and the `outlookcli` CLI.
The two surfaces expose the same 8 tools and the same 66 actions with the same parameters, so this
guidance is identical for both.

## Rule 1: Check Outlook availability before anything else

Run `application.get-status` first in a session. It reports whether classic Outlook is present.

The new Outlook for Windows has **no COM object model** and cannot be automated. If status reports
`NewOutlookOnly`, stop and tell the user they need classic Outlook. Do not retry other actions; they
will all fail.

## Rule 2: Discover, don't ask

Do not ask the user questions you can answer with a read-only call.

| Don't ask | Do this instead |
|---|---|
| "Which folder?" | `folder.list-default`, then `folder.list-children` |
| "Which message?" | `mail.read-active` for the open item, or `mail.list` / `mail.search` |
| "What is the entry ID?" | Get it from a `list`, `search`, or `read-active` result |
| "Does this message have attachments?" | `attachment.list` |

Ask the user only when the answer is a genuine preference or an irreversible decision.

## Rule 3: Confirmation gates, and why they are not everywhere

Some actions refuse unless you pass `confirm=true`. The line is drawn at **recoverability**, not at
how alarming the verb sounds. Gating a recoverable action would train you to pass `confirm=true`
reflexively, and that is exactly how the gate on the irreversible one stops being read.

**Gated - refused without `confirm=true`:**

| Action | Why there is no way back |
|---|---|
| `mail.send` | A sent message cannot be recalled, and the effect is outside the mailbox entirely. |
| `folder.delete` | Every message and every subfolder goes with the folder, and in a store with no Deleted Items it is gone outright. |
| `attachment.remove` | An attachment has no Deleted Items of its own. This destroys the only copy the message holds. |
| `calendar.delete-appointment` **with `occurrenceDate`** | Cancelling one occurrence writes a deletion exception into the recurrence pattern. Nothing lands in Deleted Items. |
| `mail.delete`, `contact.delete`, `task.delete`, `calendar.delete-appointment` — **when the item is already in Deleted Items** | There is no second recycle bin. This delete destroys the item. |

**Not gated, deliberately:**

| Action | Why it is recoverable |
|---|---|
| `mail.delete`, `contact.delete`, `task.delete` | Outlook moves the item to the store's Deleted Items folder. The user restores it from the Outlook UI. |
| `calendar.delete-appointment` (whole appointment or series) | Same: it goes to Deleted Items. |
| `mail.move` | Undone by moving the item back. The entry ID changes, so the response reports the new one. |

That these are ungated is a decision, not an oversight. It does **not** relieve you of Rule 7: say
what you deleted and where it went, so the user can undo it if you were wrong.

Two further points on `mail.send`:

- **Never send a draft the user has not seen.** Create it with `mail.create-draft`, describe the
  recipients, subject, and body back to them, and send only after they confirm.
- `mail.send` is idempotent per operation ID. If a call times out or the result is ambiguous, retry
  with the **same** operation ID rather than sending again. Generating a new ID risks a duplicate.

## Rule 3a: Rule writes need more confirmation than deletes, not less

`rule.create`, `rule.update`, `rule.set-enabled` and `rule.delete` change what happens to mail that
has not arrived yet. That makes them the highest-risk actions in this surface, above `mail.delete`:

- a message deleted in error sits in Deleted Items and the user notices within minutes
- a rule created in error silently moves or destroys **future** mail, keeps doing it, and is
  typically noticed days later

So:

- **Read before you write.** `rule.list` with `includeDetail` first, and show the user the rule you
  intend to create or change, in full, before creating or changing it.
- **Never touch a rule the user did not name.** Rules are addressed by name and Outlook allows
  duplicates; a name matching more than one rule is refused rather than guessed at, and you should
  take that refusal back to the user rather than picking one.
- **Prefer `rule.set-enabled` with `false` over `rule.delete`.** Disabling is reversible; deleting
  is not, and the user loses the definition.
- Every write rewrites the store's whole rule collection, so check `ruleCount` in the response
  against what `rule.list` reported before.

Two behaviours will otherwise look like bugs and are not:

- **A new rule is inserted first, not last.** It runs before every rule the mailbox already had.
- **`deleteMessage` reads back as `moveToFolder`.** Outlook has no delete action; it stores
  "delete it" as a move to Deleted Items plus stop-processing.

There is no mark-as-read rule action, in either surface. Outlook's rule object model does not have
one, so this is not something to work around - tell the user only the Rules and Alerts wizard in
Outlook can do it.

**Note that no `rule` action is gated behind `confirm=true` today**, so its absence from Rule 3's
table is not a statement that rule writes are recoverable. By Rule 3's own recoverability line
`rule.delete` is not - the definition is gone and there is no Deleted Items for it - so the
confirmation here is yours to obtain from the user, not the API's to enforce. Treat this section as
carrying that weight until the gate exists.

## Rule 4: Entry IDs are the addressing scheme

Outlook items are addressed by entry ID, not by name or index. Entry IDs:

- come from a `list`, `search`, or `read-active` result
- are stable for an item in a folder, but **change when the item is moved between stores**

So re-read after a move rather than reusing the old ID.

## Rule 5: Drafts before edits

`set-subject`, `set-body`, `set-recipients`, `attachment.add`, and `attachment.remove` target drafts.
Create or open a draft first, apply the edits, then send. Do not attempt to edit an already-sent
message.

## Rule 6: Prefer narrow queries

`mail.list` and `mail.search` can return very large result sets from a busy mailbox. Constrain by
folder and by count. Read full bodies with `mail.read` only for the items you actually need, rather
than pulling bodies for an entire folder.

A single call returns one page. If the response says `hasMore: true`, do not report the result as
complete - either page through with `cursor` until `hasMore` is false, or say plainly that you only
looked at part of the folder. "I found no such mail" is wrong when there was more you did not read.

## Rule 7: Report what you did, in the user's terms

Never leave a tool call as the entire response.

| Bad | Good |
|---|---|
| *(tool call, no text)* | "Found 12 unread messages in Inbox; 3 are from Finance." |
| "Done." | "Created a draft to alice@example.com, subject 'Q1 numbers'. Not sent - confirm and I'll send it." |
| "Deleted." | "Moved 4 newsletters to Deleted Items." |

State recipients and subject explicitly before any send. That is the user's last chance to catch a
mistake.

## Rule 8: Read the error, then act on it

Errors return structured, actionable context:

```json
{
  "success": false,
  "errorMessage": "Mail item not found for the supplied entry ID",
  "suggestedNextActions": ["mail.list", "folder.list-default"]
}
```

- `success: false` always accompanies an `errorMessage`. Never treat a result with an error message
  as a success.
- Follow `suggestedNextActions` before improvising.
- If the error says Outlook is unavailable, go back to Rule 1 rather than retrying in a loop.

## Rule 9: Outlook is a shared desktop application

Outlook is a single running instance the user is also using. Automation is not isolated:

- the user may be typing in a window you are editing
- a modal dialog in Outlook can block calls
- operations run on the user's real mailbox, with real consequences

Work in small, verifiable steps and re-read state rather than assuming your last write stuck.
