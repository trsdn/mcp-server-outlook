# Changelog

All notable changes to OutlookMcp, an Outlook COM automation MCP server, will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added

- **Address book lookup and recipient resolution** (#15): a new `addressbook` tool with `resolve`,
  `list-address-lists` and `list-entries`. `Session.AddressLists` was previously unused, so there
  was no way to check an addressee before sending to them - the only feedback available was a bounce.

  `resolve` takes one or more display names, aliases or addresses and answers per addressee, with
  `allResolved` as the single flag to check before sending and `unresolvedNames` naming the bad
  ones. An unresolved name is a success with `resolved: false`, not a failure: "no such person" and
  "Outlook could not be reached" are different answers and collapsing them would defeat the point of
  the call.

  Three things were measured against a live Exchange mailbox rather than assumed:

  - **`AddressEntry.Address` is not an email address.** For an Exchange entry it is the X500
    legacyExchangeDN (`/o=.../cn=Recipients/cn=...`) - a plausible-looking string that mail cannot be
    sent to. The real address comes from `GetExchangeUser().PrimarySmtpAddress`, with
    `GetExchangeDistributionList().PrimarySmtpAddress` for a group (a different call on a different
    COM type; `GetExchangeUser()` returns null for a group) and a `PR_SMTP_ADDRESS` read as a
    fallback for everything else, including any entry at all while the client is offline from the
    directory. Both values are reported - `smtpAddress` and `rawAddress` - and `smtpAddressSource`
    says which route produced the answer, so a directory-confirmed address is distinguishable from a
    one-off SMTP string that was never checked against anything.

  - **An ambiguous name is indistinguishable from a missing one.** `Recipient.Resolve()` returns
    false for both and the object model offers no way to list the candidates, so both are reported as
    unresolved rather than guessed at.

  - **An address book cannot be searched, only scanned.** There is no `Restrict` or `Find` on
    `AddressEntries`, so `list-entries` walks the book with the documented `GetFirst`/`GetNext`
    cursor and stops at `scanLimit`. `scanLimitReached` distinguishes "the book ran out" from "the
    scan did", because an empty answer means opposite things in the two cases.

  This is the most Object Model Guard-exposed surface in the product: every property and method on
  `Recipient` is protected, along with `AddressEntry.Address`, `GetExchangeUser`,
  `ExchangeUser.PrimarySmtpAddress` and every cursor method on `AddressEntries`. Guard exposure is
  documented on every action, a whole-call denial fails with an explanation rather than a wrong
  answer, and a property refused while the rest of the call succeeded is named in `accessDenied` so a
  missing value is never confused with a value the directory does not hold.

  `IsObjectModelGuardDenial` was deliberately left testing `E_ABORT` only. Microsoft documents the
  guard returning `MAPI_E_NOT_SUPPORTED` (0x80040102) for a protected member refused outright, so
  widening it was tried - and reverted, because 0x80040102 is also the ordinary MAPI "this provider
  or property type does not support that" error. Treating it as a denial would make every such
  failure across the existing mail and folder surface tell the caller to look for a security dialog
  that does not exist.

### Added

- **Room and equipment booking** (#32): `calendar create-appointment` takes `resourceAttendees`
  alongside `requiredAttendees` and `optionalAttendees`. Previously a room could not be booked at
  all - an agent asked to "book a room" could only name it in `requiredAttendees`, which invites the
  room like a person and reserves nothing, or put it in `location`, which is a free-text label. The
  read side already reported `olResource` recipients as `"resource"`, so this closes the write half
  of a gap that was only ever half implemented.

  Two things were measured against a live mailbox rather than assumed:

  - **`AppointmentItem.Resources` does not work through the embedded interop types.** Assigning it
    attached no recipient at all, while the identical late-bound assignment attached one every time.
    `RequiredAttendees` and `OptionalAttendees` are unaffected, so the failure is silent and specific
    to resources: the meeting saves, reports success, and has no room on it. Resources are now added
    as recipients explicitly and typed `olResource`, which does not depend on that property.

    This was caught only by exercising both surfaces. The in-process test went green first, and the
    CLI run against the same code returned `"attendees": []`. Another instance of the project's
    characteristic failure mode, and the first one that a single-surface check would have shipped.

  - **Resolution is a weak existence check.** Anything SMTP-shaped resolves as a one-off external
    address whether or not the mailbox exists, so an invented room address is accepted. Only a bare
    name absent from the address book fails to resolve. The refusal path still guards that case, and
    the limit is now documented rather than left to be discovered.

- **Item export** (#14): `mail export` and `calendar export` save an item to disk via `SaveAs`.
  Mail writes `.msg`, `.txt`, `.html`, `.mht` or `.rtf`; appointments add `.ics`. This was the last
  unchecked acceptance criterion on the item-coverage epic.

  `SaveAs` looks trivial and is not. Four behaviours were measured against real Outlook first, and
  each one hands a caller a confidently wrong result if passed straight through:

  - **The ANSI `.msg` format silently destroys text.** A subject reading
    `probe Grüsse äöü тест €` saved with `olMSG` and reopened comes back as
    `probe Grüsse äöü ???? €` - no error, no warning. `olMSGUnicode` round-trips it exactly, so
    `msg` always means the Unicode variant here and the ANSI constant is never used. Note that
    German umlauts and `€` survive ANSI on a CP1252 machine, so a careless test would miss this
    entirely; only characters outside the machine's code page are lost.
  - **The extension is ignored.** `SaveAs("x.txt", olMSGUnicode)` succeeds and writes an OLE
    compound file under a `.txt` name. A `format` that contradicts the extension is now refused,
    without writing anything.
  - **A relative path succeeds and lands elsewhere.** Outlook resolves it against its own working
    directory rather than the caller's, so the file exists somewhere nobody will look. `filePath`
    must now be absolute.
  - **A missing directory reports "The operation failed."**, which names nothing. The directory is
    now checked and named before any COM object is touched.

  Overwriting is also silent in Outlook, so an existing file is never replaced unless `overwrite`
  is passed, and the result reports whether it was.

  Verified live against classic Outlook: 12 integration tests, none skipped, plus a round trip
  through the CLI and through a real MCP stdio session. The Unicode mapping and the overwrite guard
  were both sabotage-proved - reverting `msg` to `olMSG` fails the round-trip assertion on the
  Cyrillic, and disabling the guard fails the overwrite test.

- **`task` tool** (#14): `list`, `read`, `create`, `update` and `delete` over Outlook's Tasks folder,
  the last item type in the mailbox with no coverage at all. Available identically through the MCP
  server and the CLI.

  Two behaviours were measured against a real Tasks folder before being designed around, and both
  would have made the listing actively misleading:

  - **Outlook stores 1 January 4501 for a date that was never set** - not null, not the OLE zero
    date. On the folder this was built against that sentinel is 260 of 274 due dates, so passing it
    through would have dated 95% of a listing to the 46th century. Unset dates are now omitted.
  - **271 of those 274 tasks were already complete**, so `list` omits completed tasks by default.
    That filter reports itself: `completedItemCount` counts what was hidden, `includedCompleted`
    echoes the flag, and `scannedItemCount` always equals the returned rows plus the skipped ones, so
    a filtered listing can never be mistaken for an empty folder.

  Both behaviours were confirmed load-bearing by removing them and watching the corresponding test
  fail.

- **`application get-active-explorer` and `application get-active-inspector`** (#14): report which
  folder the user is looking at, what is selected there, and which item they currently have open.
  `get-active-explorer` returns the current folder's full store path, so an agent can address the
  folder the user is actually in instead of guessing at a folder name.

  Both report "nothing is open" as a success rather than an error, since no explorer and no open
  inspector are ordinary states.

  The item kind is derived from a typed switch over the Outlook interfaces. `GetType().Name` on an
  Outlook item returns `__ComObject`, so the obvious implementation would have labelled every item
  in the mailbox `"__ComObject"`.

  An item the user is still composing has not been saved and has no `entryId`, so it cannot be
  addressed by any other action; the field is omitted and `isSaved` distinguishes that case from a
  saved item. Note that such an item's parent folder reports as the Outbox - where it *would* be
  sent - not the Drafts folder.

### Fixed

- **The test suite intermittently crashed the test host** (#116). Thirteen sites across ten
  integration test files released the shared `Outlook.Application` with
  `OutlookInteropRunner.ReleaseComObject`, whose generic overload is `FinalReleaseComObject`. That
  zeroes the RCW refcount for every holder in the process, so a later test could be handed a wrapper
  separated from its RCW and the host died with `STATUS_STACK_BUFFER_OVERRUN` (`0xc0000409`). It
  never surfaced as a test failure, only as apparent infrastructure flakiness. This is the same
  defect as #19, which production already guards against; the convention was documented in the test
  suite but not followed. All sites now use `ReleaseSharedComObject`.

  A new pre-commit check, `check-shared-application-release.ps1`, keeps it from coming back. It was
  confirmed to block a commit by reintroducing one of the thirteen sites.
- **The `contact` tool was never exposed over MCP.** It was generated, implemented, routed through
  the CLI, counted in README.md and FEATURES.md, and documented as being "identical through the MCP
  server and the CLI" - while missing from the explicit tool allow-list in `Program.cs`, so it never
  appeared in `tools/list` and no MCP client could call it. It has been absent for its entire life.

  Nothing caught it because both sides of the check were hand-maintained copies: the allow-list in
  `Program.cs` listed five tools, and the end-to-end `tools/list` test asserted against its own
  separate hand-written list of the same five. The test agreed with the bug.

  Both are now derived from one source. `Program.RegisteredToolTypes` is the single list, the
  end-to-end test registers that object rather than a copy of it and derives its expectations from
  the `[McpServerTool(Name = ...)]` attributes on it, and a new `ToolRegistrationTests` asserts that
  every generated `[McpServerToolType]` in the assembly is in it. A tool that is generated but not
  registered now fails the build.

  Found only by driving a real MCP client over stdio and reading `tools/list` - the same lesson as
  #81 and #82: a green suite said nothing, one live call said everything.

- **Registering the tool list by reference silently registered no tools at all.** While fixing the
  above, passing the allow-list as `Type[]` bound to the generic
  `WithTools<TToolType>(builder, TToolType instance, ...)` overload by exact inference instead of the
  intended `WithTools(IEnumerable<Type>)`, so the server started cleanly, advertised no tools
  capability, and answered `tools/list` with "Method 'tools/list' is not available". No compiler
  error, no startup error. The list is typed `IEnumerable<Type>` so the non-generic overload wins,
  and five tests now fail if that binding regresses.

- **Pre-commit staged the generated skill files before the build that generates them.** The
  auto-staging step asserted in a comment that "the Release build already ran"; it had not - the
  only Release build in the hook is the one inside the CLI workflow smoke test, which runs
  afterwards. So editing `skills/shared/*.md` staged the source while the generated copies under
  `skills/outlook-*/references/` were still stale on disk, and the build that refreshed them
  happened too late for them to be included in the commit. The hook now runs the Release build
  first and fails the commit if it does not succeed, so the claim holds by construction.

### Removed

- **85 dead PowerPoint result types** deleted from `ResultTypes.cs` (123 down to 38), along with the
  18 tests that were the only thing still referencing them. Nothing in the product could reach any
  of them; they were inherited from the fork and survived both cleanup epics because the earlier
  sweep matched on keywords like `Slide` and `Shape` rather than checking every type by name.

- **88 orphaned PowerPoint section headers** left behind in `ResultTypes.cs` by that deletion. Every
  one was an empty banner comment - `// -- Chart --` and similar - with no type under it.

### Changed

- **`FEATURES.md` and `README.md` now match the generated surface.** Both had drifted badly: they
  claimed 5 tools and 30 operations while the registry had 6 tools and 48. `mail` was listed with 16
  of its 22 actions and `folder` with 4 of its 10, so `get-conversation`, `respond-to-meeting`,
  `set-flag`, `list-categories`, `list-rules`, `list-reminders`, `get-free-busy`, `list-stores`,
  `open-shared` and the folder mutation actions were all shipped but undocumented.

- **`ResultTypeInvariantTests` is now driven by reflection** over every model type instead of 24
  hand-written per-type copies. The old tests asserted the same two invariants once each, every one
  of them against a PowerPoint type, so no Outlook result type had ever been checked. The new
  version covers all of them and every type added later, and it distinguishes non-nullable
  collections (which must default to empty) from deliberately nullable ones such as
  `MailSummaryInfo.AccessDenied`, where null means "nothing was blocked" and is omitted from the
  wire.

### Added

- **`contact` tool: list, read, create, update and delete Outlook contacts** (#14). Recovers the
  last unmerged slice of the orphaned `feature/outlook-parity-slices` branch. A Contacts folder
  holds distribution lists as well as people, and the original stranded implementation cast every
  item to `ContactItem` and skipped whatever did not cast - so on a real mailbox it silently dropped
  a distribution list while still reporting the full folder item count. Distribution lists are now
  returned separately in `distributionLists` with their member count, and `contacts`,
  `distributionLists` and `skippedItemCount` together always account for every item scanned, so a
  silent drop cannot recur. `update` writes only the fields it is passed, leaving the rest alone.
  Implemented with typed PIA calls throughout - no `dynamic` anywhere in the surface.

- **`mail get-conversation` now names the non-mail members of a thread** (#111). Meeting
  invitations, the calendar appointments they create, and acceptances or declines are returned in a
  new `otherItems` array with a named type, subject, folder and timestamp, instead of being reduced
  to a number. On the thread this was measured against, four of seven items were previously
  invisible - including the invitation and the acceptance, which are usually the substance of the
  conversation. `messages` is unchanged and still mail-only, so nothing that reads it today breaks.
- **`skippedItemCount` now means only "could not be read"** (#111). It previously also counted every
  meeting item, so a thread read perfectly and a thread with unreachable entries reported the same
  non-zero number and neither was actionable.
- **`mail list-reminders`** (#15). Lists what Outlook intends to remind the user about - appointments,
  tasks and flagged mail together - earliest first, with counts for the whole set.

  Three behaviours were measured against a real mailbox, and each is a trap a plausible
  implementation walks into.

  The due time comes from `OriginalReminderDate`. `NextReminderDate` is the obvious-looking choice
  and is only populated once a reminder has snoozed or recurred: 152 of the 605 reminders on the test
  mailbox sit at the OLE zero date, so a quarter of the listing would arrive dated 1899.

  Results are sorted here, because Outlook does not return the collection in date order. Applying a
  limit to its native order hands back an arbitrary handful out of six years of reminders.

  Overdue reminders are excluded by default — 416 of 605 on the test mailbox, the oldest five years
  old — because including them buries everything still to come. `overdueCount` is always reported so
  a caller knows what was held back rather than mistaking a page for the whole picture.

  `IsVisible` is deliberately not surfaced: it means "the reminder dialog is on screen right now",
  which was false for all 605, so presenting it as pending-ness would report a mailbox full of
  reminders as having none.

  Read-only; dismissing and snoozing are not exposed.

- **`mail list-rules`** (#15). Enumerates the mailbox's inbox rules, optionally with each rule's
  conditions, actions and move-to destination.

  Rules move, delete and forward mail before anything else in this surface sees it. Until now they
  were invisible, so "why is nothing arriving from this sender?" got a confident empty folder instead
  of the answer. That is the project's characteristic failure mode: a truthful response to a question
  the caller did not ask.

  Two behaviours were measured against Outlook rather than assumed. `Conditions` and `Actions` are
  **fixed-length collections** covering every clause Outlook supports — every rule on the test mailbox
  reported 31 conditions and 28 actions regardless of content — so only the clauses with `Enabled`
  set are reported. And rule recipients are stored **unresolved**: `Recipient.Address` is blank and
  the address sits in `Name`, so reading only `Address` would report every from-rule as matching
  nobody.

  `includeDetail` is off by default because walking those slots costs roughly forty times as much:
  on a mailbox with 84 rules, 264 ms became 10.6 s. Clause names and rule types are reported as names,
  never as raw enum ordinals.

  Read-only. Creating or changing a rule alters real mail flow for every future message and is
  deliberately not exposed.

- **Follow-up flags: `mail set-flag`** (#15). Raise, complete or clear a follow-up flag, with an
  optional due date and label. `read`, `list` and `search` now report `flagStatus` — always, as
  `none`, `flagged` or `complete`, never omitted — plus `flagRequest` and `flagDueDate` when set, so
  "what still needs following up?" is one listing call rather than one read per message.

  `complete` and `none` are deliberately distinct. Collapsing them would report work that was done as
  work that was never raised.

  Two behaviours were measured against Outlook rather than assumed, and both would otherwise have
  shipped as silent wrong answers. `MarkAsTask` is refused on drafts (raised as
  `NotImplementedException`, not `COMException`, so the obvious catch misses it) and completing a flag
  on a draft is refused outright — that now returns an explanation instead of a raw COM error. And
  `ClearTaskFlag()` reports success on a draft while leaving the flag exactly where it was; clearing
  assigns the state directly and also resets the task dates, which otherwise survive and surface as a
  due date on an unflagged message.

  An unrecognised `flagStatus` is refused before the item is touched, so nothing is half-applied.

- **`mail list-categories`** (#15). Enumerates the mailbox's master category list, so a caller can
  discover which categories exist before writing one.

  This closes a real hole in the already-shipped `set-categories`. Outlook does not validate the
  string it writes: a name that is not in the list is accepted, reported as `success: true`, and only
  later turns out to be a label that cannot be filtered or coloured by. A typo did not fail, it
  quietly created dead data — the project's characteristic failure mode again.

  Colours come back as names (`yellow`, `darkTeal`, or `none`), never as the raw `OlCategoryColor`
  ordinal, because a number is not something a model can show a user or reason about. The mapping is
  derived from the enum rather than a hand-written table of 26 entries, which would report confidently
  wrong colours as soon as it drifted. A category with no name is skipped rather than listed, since it
  could not be passed back to `set-categories` anyway.

- **`flaggedOnly` filter on `mail list` / `mail search`** (#15, #27). Returns only messages with an
  outstanding follow-up flag. Pushed into Outlook as a DASL `Restrict` clause on
  `PR_FLAG_STATUS` — which has no `urn:schemas:httpmail:` equivalent and so is addressed by MAPI
  proptag `0x10900003` — rather than hydrating the folder and filtering afterwards. Verified against
  a live mailbox by comparing `scannedCount` with and without the filter, since a client-side
  implementation would return identical results while doing all the work.

  The clause is `= 2`, deliberately not `<> 0`. A completed flag is finished work, and returning it
  under "flagged" would put items the user has already dealt with back on their to-do list. The same
  narrowing is applied client-side after `Restrict`, because the DASL filter is over-inclusive by
  design and a caller can reach that path with nothing pushed down. Meeting requests carry the same
  flag and are filtered identically instead of being dropped for being a different item type, and
  they now report `flagStatus` in listings so a returned item cannot claim to be unflagged.

  `flaggedOnly` is part of the paging cursor's fingerprint, so a cursor minted under one filter is
  not silently accepted under another.

- **HTML message bodies** (#15). `create-draft`, `reply`, `reply-all`, `forward` and `set-body` take
  `bodyFormat`, which is `plain` (the default, and what every existing caller keeps getting) or
  `html`. Composing a bulleted list, a link or a table no longer means writing tags and watching them
  arrive as visible tag soup.

  Plain text stays escaped rather than interpreted, which matters most when the text came from the
  user: `profit < loss` sent as HTML loses everything after the bracket to what the renderer takes
  for an unclosed tag, and the send still reports success. An unrecognised `bodyFormat` is refused
  rather than quietly treated as plain.

- **Folder mutation: `folder create`, `rename`, `move` and `delete`** (#15). The `folder` tool could
  only read, so "file these into a 2024 folder" ended at the first step.

  **Default folders and store roots are refused for rename, move and delete.** Outlook itself
  permits deleting the Inbox - no prompt, no error, and every message in it goes too - so this guard
  is the substance of the feature rather than a nicety around it. It compares entry ids, not names,
  across *every* store in the profile: a folder merely called "Inbox" that is not the default Inbox
  stays deletable, and an archive's Deleted Items is protected as much as the primary mailbox's.

  Also refused: creating a folder with a blank name or one containing a backslash (both produce a
  folder that cannot afterwards be addressed by path), creating a duplicate sibling name (the
  existing folder is deliberately *not* returned - a caller expecting a new empty folder would
  otherwise be handed one with contents in it), and moving a folder into its own subtree.

  Emptying a folder in place is deliberately not included; delete the items or delete the folder.

  Found while testing, and fixed here: **setting `MAPIFolder.Name` does not refresh the reference**,
  so reading the name and path back from it returns the old values. A rename therefore reported
  success while appearing not to have happened. The folder is now re-fetched by entry id, which
  Outlook preserves across a rename.

  That same staleness left four real folders in the developer's mailbox: the tests cleaned up by the
  path the operation had returned, and when that path was stale the delete quietly failed with its
  result discarded. They were found by listing the Inbox afterwards, not by any assertion. Cleanup
  now sweeps by name prefix, twice, and reports what it could not remove.

- **Shared and delegate mailboxes** (#38). New `folder open-shared --address <smtp> --role <role>`,
  where `role` is one of `inbox`, `calendar`, `contacts`, `tasks`, `notes` or `journal`. The returned
  folder path works with `mail list`, `calendar list` and `folder list-children` like any other.

  Previously the only mailboxes reachable were those already in the profile, so "read the team
  inbox" or "check a colleague's calendar" had no answer at all.

  **`address` is required, and an address that cannot be opened is refused.** This is the whole
  reason the operation is written the way it is: `Recipient.Resolve` returns `false` rather than
  throwing, and Outlook then treats an unresolved recipient as the current user - so a defaulted or
  mistyped address would return *your own* inbox with `success: true`. Silently answering about the
  wrong mailbox is worse than any error.

  Verified live that the resolve guard alone is **not sufficient**: `Resolve` returns `true` for
  `no-such-person@invalid.example`, because Outlook accepts any syntactically valid SMTP address as
  a one-off recipient without consulting the directory. The refusal in that case comes from
  `GetSharedDefaultFolder` failing, which is why the error names the mailbox and the role and states
  plainly that Outlook cannot distinguish "no such mailbox" from "no access granted".

  Tested against the signed-in user's own address, which is the only mailbox a test can assume
  access to - `open-shared` for that address returns the same folder path as `list-default`, which
  is what makes the test non-vacuous. **Access to a genuinely foreign mailbox is unverified**; this
  profile has no delegate rights to borrow.

- **Store discovery, and default folders from a specific mailbox** (#38). New `folder list-stores`,
  and `folder list-default` takes `storeId`.

  Everything targeted the default delivery store implicitly. `NameSpace.GetDefaultFolder` always
  reads that one store, so `folder list-default` reported a single mailbox's folders and said nothing
  about the others existing - a caller with an archive or a second account was told, with
  `success: true`, that its Inbox was the only Inbox. Verified on the developer's own profile: it has
  two stores, and the Online Archive was unreachable through the tool.

  `list-stores` reports `storeId`, `displayName`, `isDefaultStore`, `isDataFileStore`,
  `exchangeStoreType`, `filePath`, `rootFolderPath`, and the delivering account's address where one
  exists. Accounts are folded into the store list rather than exposed separately, so a caller never
  has to correlate two lists by id to answer "which mailbox is this?".

  Folder results now carry `storeId` and `storeName`, because on a multi-store profile a folder
  listing that does not name its store is ambiguous and the caller has no way to notice.

  An unknown `storeId` is **refused**, not quietly served from the default store. Returning real
  folders and real item counts from a mailbox the caller did not ask for is the worst outcome
  available here. A store that genuinely lacks a default role reports `available: false` with a
  `note` saying so, rather than failing the call.

### Fixed

- **The pre-commit hook silently dropped generated skill files from the commit.** Its
  auto-staging step ran `git diff 2>&1`; git writes advisory notices such as "LF will be replaced
  by CRLF" to stderr, and under Windows PowerShell with `$ErrorActionPreference = 'Stop'` that
  notice became a terminating error. `git add` never ran, the `catch` printed
  "Continuing with remaining checks...", and the script went on to report
  "All pre-commit checks passed!" - so a commit could ship a changed tool surface with the
  regenerated `SKILL.md` and reference docs left behind. The block now relaxes the preference
  around the git calls, checks `$LASTEXITCODE`, re-queries to prove the files really are staged,
  and fails the commit instead of continuing. Same shape as #82: a check reporting success without
  having done its job, and again only visible under `powershell.exe`.

- **Replying flattened the quoted thread to plain text** (#15). Passing a `body` to `reply`,
  `reply-all` or `forward` read the draft's plain-text `Body` - a lossy projection of a quoted
  original that is almost always HTML - and wrote it straight back. Every word survived, so the call
  reported success and nothing looked wrong until someone opened the draft and found the original's
  tables, inline images, links and emphasis gone. The recipient saw a flattened thread.

  The caller's text is now inserted into `HTMLBody`, just inside the `<body>` element so it sits
  above the quote, and the quoted original is left untouched. Plain text is HTML-escaped on the way
  in, because otherwise a user writing `profit < loss` would silently lose the rest of the sentence.

  Worth recording for anyone testing this area: the obvious assertion - "the reply is still
  `BodyFormat=HTML`" - passes against the bug, because writing plain text into an HTML draft leaves
  the format alone and makes Outlook regenerate a wrapper. So does checking that links survived:
  **Outlook auto-linkifies bare URLs** when it converts plain text to HTML, so `<a href>` reappears
  in a body that has just been stripped of everything else. Both were measured by deliberately
  reinstating the bug and watching the tests stay green. The assertion that actually holds is that
  markup Outlook *cannot* invent - `<img>`, `<table>`, `<b>`, `<li>` - survives from the original
  into the reply.

- **Folder paths resolved against a stale listing** (#15). Resolving `\\store\Inbox\Project` walked
  `NameSpace.Folders` recursively, comparing each child's `Name`. Two problems, one of which is a
  correctness bug rather than a slow path.

  Outlook's `Folders` enumeration goes **stale after a rename and stays stale for the life of the
  process**: the enumerated child keeps reporting its old name indefinitely. So a folder renamed
  through this tool could not then be addressed by the path the rename itself had just returned - the
  caller got `Unsupported Outlook folder` for a folder that plainly existed. Cross-process it worked,
  which is what made it look like a timing lag; it is not. Bounded retries were tried and failed at
  the same rate, which is what finally identified the cache.

  Path lookup now walks segment by segment through the `Folders["name"]` indexer, which asks Outlook
  instead of reading a cached listing. It sees renames immediately, and it is O(depth) rather than a
  depth-first scan of every folder in every store - the folder test suite went from 50s to 5s.
  Enumeration is still used for bare names and store-less paths, where there is no segment to walk.

- **A folder path that was not a path** (#38). `GetFolderPath` returned whatever Outlook's
  `FolderPath` gave it. For a folder that is not in a store's tree, that is the folder's **entry
  id** - a long hex string that looks like a value, is accepted everywhere a path is, and resolves
  to nothing.

  This surfaced the moment store targeting made archives reachable. `Store.GetDefaultFolder` answers
  for every default role whether or not the store has one: the developer's online archive contains
  four folders, but reported all ten roles as available, nine of them with an entry id where the
  path should be. A caller would have passed one straight back to `mail.list` and been told the
  folder was empty.

  A real Outlook folder path always begins with `\\`, so anything else is now reported as absent.
  This is a general fix in the shared helper, so every result that carries a folder path benefits,
  not just the new store paths.

  Found by running the tool against a live mailbox. The first version of the new tests asserted "10
  of 10 roles available" and **passed** - a check reporting success without having verified
  anything, which is the failure mode this project keeps rediscovering. They now assert that an
  available role's path round-trips through `resolve-path`.

- **Full-text search via Outlook's content index** (#42). `mail.search` takes `searchMode`, and every
  search response now reports `searchEngine`.

  The default free-text path opens each candidate message and does a substring check. That is exact,
  but it is bounded by a scan limit, so in a folder larger than that limit a genuine match further
  back is never reached and the caller is told - with `success: true` - that no such mail exists.
  `searchMode: "fullText"` pushes the query down to Outlook's content index instead: nothing is
  opened client-side and there is no scan horizon.

  The two engines are **not** fast and slow versions of the same question. The index matches whole
  words; the scan matches substrings, so it finds `foo` inside `foobar` and the index does not. That
  is why the mode is opt-in and why `searchEngine` (`clientScan` or `contentIndex`) is on every
  response: an empty result means different things depending on which engine produced it, and a
  caller cannot tell them apart otherwise. If the index was asked for and the store could not serve
  it, `searchEngine` says `clientScan` and `message` explains why, rather than quietly handing back
  substring semantics under the label that was requested.

  The pushed-down clause spans body, subject, sender name, sender address, To and Cc - the same
  fields the client-side check reads. Pushing only the body down would be under-inclusive in the
  worst way: Outlook would discard a subject match before the client ever saw the item.

  A page cursor is bound to the engine as well as the query, so one minted by the scan cannot be
  replayed against the index.

  Verified against classic Outlook: the pushdown is evidenced by direct comparison - for a term
  nothing matches, the client scan examined every item in the folder and the index examined none.

- **Single-occurrence changes and cancellations** (#33). `calendar.update-appointment` and
  `calendar.delete-appointment` take an optional `occurrenceDate` that limits the change to one
  instance of a recurring series.

  The trap this closes is silent. Every occurrence of a series carries the *master's* entry id, so
  updating or deleting "an occurrence" by its entry id changed or cancelled the **whole series**
  while looking exactly like a single-instance edit. Cancelling one stand-up wiped out every
  stand-up, and nothing reported an error, because as far as Outlook was concerned that is what was
  asked for.

  So the scope of the change is now something the caller states rather than something they discover
  afterwards. The response reports `scope` as `series` or `occurrence`; naming `occurrenceDate` on an
  item that is not recurring is refused rather than ignored; and a date the series does not fall on
  is refused rather than being quietly turned into a series-wide edit.

  `occurrenceDate` may be a bare date. Outlook's `GetOccurrence` matches on the occurrence's exact
  start to the minute and throws if it is off by one, so a value with no time of day takes its time
  from the series - a caller asking to cancel Thursday's stand-up should not have to know it starts
  at 09:17.

  Verified against classic Outlook: 6 new integration tests, confirmed failing first. The two
  load-bearing ones assert that the *rest of the series survives*, not merely that the named instance
  is gone - asserting only the absence would pass just as happily if everything had been wiped out,
  which is the bug.

- **Recurring appointments** (#33). `calendar.list` now expands a recurring series into its
  individual occurrences, and `calendar.create-appointment` can create one.

  This closes a hole that was worse than a missing feature. Outlook stores a series as a single
  master item dated at its first occurrence, so a listing that does not ask for expansion returns
  *nothing* for a weekly stand-up when you ask about next Tuesday - and returns it as a confident
  empty list. An agent asked "am I free Tuesday at 10?" would have said yes.

  Expansion needs both `start` and `endTime`, because a series with no end date has infinitely many
  occurrences. The result reports `recurringExpanded` so a caller can tell whether the list is
  complete, and says plainly when it is not. Every listed item now carries `recurrenceState`
  (`notRecurring`, `master`, `occurrence`, `exception`), and `calendar.read` reports the stored
  pattern including how many occurrences deviate from it.

  Series creation takes `recurrenceType` (`daily`, `weekly`, `monthly`, `yearly`),
  `recurrenceInterval`, `recurrenceDaysOfWeek`, and either `recurrenceCount` or `recurrenceEndDate`.
  An unusable pattern is refused before anything is written, rather than quietly producing a single
  appointment and reporting success.

  Verified live: a five-day daily series created and listed back as five occurrences on five distinct
  dates.

  One trap worth recording, found only by running against a real mailbox. Outlook's Jet restriction
  syntax (`[Start] <= '...'`) parses its date literal using the **machine's regional settings**, not
  a fixed format. A US-formatted literal on an en-DE machine had its day and month swapped and
  matched 2 appointments where the correctly formatted one matched 12. It fails asymmetrically -
  a whole-day window still works, because midnight survives a mangled time - so it only breaks the
  intraday questions where a wrong answer matters most. Note this is the *opposite* of the DASL
  `@SQL=` filters used for mail, which take a culture-independent UTC literal.

- **Answering meeting invitations** (#32). `mail.respond-to-meeting` accepts, declines or tentatively
  accepts an invitation. Previously the only thing an agent could do with one was reply to it, which
  is just mail back to the organiser and leaves the invitation unanswered - a confidently wrong
  answer of exactly the kind this surface exists to avoid.

  Answering and notifying are separate. The response always updates the caller's own calendar; the
  organiser is mailed only when `sendResponse` is set, which defaults to false because mail to a real
  organiser cannot be taken back. `responseText` adds a note to that mail.

  Anything that is not an invitation is refused by name rather than failing obscurely inside Outlook:
  ordinary mail, a meeting cancellation (the meeting is already off) and somebody else's response to
  a meeting the caller organised.

  Verified live only along the refusal paths - an unknown response value, no item, ordinary mail and
  a meeting response are each rejected against a real mailbox. **The accept path is unverified**, and
  deliberately so: there is no way to manufacture a safe invitation, since you cannot invite yourself
  and then answer it, and every real invitation in this mailbox belongs to a real organiser. An
  on-demand test (`OUTLOOKMCP_RESPOND_ENTRYID`) exists so the owner can verify it against an
  invitation they are willing to answer.

- **Free/busy lookup** (#32). `calendar.get-free-busy` answers "when is this person available",
  which was previously impossible - nothing exposed `Recipient.FreeBusy`, so an agent asked to
  schedule around somebody could only guess. It returns Outlook's raw slot string and, more usefully,
  the non-free stretches decoded into timestamps with a status (`tentative`, `busy`, `outOfOffice`,
  `workingElsewhere`).

  Two deliberate refusals to overstate the answer. Outlook ignores the requested length and returns
  however far ahead it publishes, so the result is trimmed to the window asked for and `end` is
  pulled *in* when Outlook returned less - with a `message` saying so, because padding it out would
  invent free time nobody looked up. And an attendee Outlook cannot resolve fails the call: Outlook
  reports an all-free calendar for a recipient it never looked up, so treating that as an answer
  would schedule over a calendar nobody read.

  Verified against classic Outlook: a real query returned 136 busy periods across the published
  window, matching the owner's actual calendar.

- **Attendees on calendar items** (#32). `calendar.create-appointment` accepts `requiredAttendees` and
  `optionalAttendees` (semicolon-separated), which turns the item into a meeting, and
  `sendInvitation` mails them. `calendar.read` reports `isMeeting` and an `attendees` list with each
  person's type and `responseStatus`. Previously the calendar surface could only ever produce a solo
  appointment: an agent asked to "set up a call with Anna" would put an entry in the caller's own
  calendar, report success, and never tell Anna.

  Creating a meeting and inviting people are deliberately separate. The meeting is saved to the
  caller's own calendar and nobody is told until `sendInvitation` is set, and the response message
  says so in words rather than leaving it to be inferred.

  An attendee Outlook cannot resolve is a **failure**, not a warning: the meeting is not created and
  `unresolvedAttendees` names them. Outlook will happily save such a meeting, so the naive behaviour
  is to report success for a meeting that can never reach the person the caller named.

  Verified against classic Outlook, including a real self-addressed invitation under
  `RunType=OnDemand` with the self-addressing constraint enforced in code. Delivery itself is *not*
  verified - with the owner as the only attendee there is nobody for Exchange to notify - and the
  test says so rather than implying more than it checked.

- **Conversations and threads** (#39). `mail.get-conversation` returns a whole thread from any one
  message in it, ordered oldest-first and spanning folders, so a reply sitting in Sent Items comes
  back alongside the original in the Inbox. `mail.read`, `mail.list` and `mail.search` now carry
  `conversationId` and `conversationTopic`, and listings carry `folderPath`, so reaching a thread
  costs one call rather than a read per message.

  Each item reports the folder it lives in, because knowing a thread exists is useless if you cannot
  act on its parts. Items that are not mail - a meeting response in the middle of a thread, say - are
  counted in `skippedItemCount` rather than dropped, so the count you are shown adds up.

  A store with conversation view disabled returns an explicit failure carrying
  `conversationSupported: false`, not an empty success. "This message has no replies" and "I cannot
  tell you whether this message has replies" are different answers and must not look alike.

  One caveat, found by running the tests rather than by reading the API: Outlook's conversation index
  is eventually consistent. A reply created a moment ago is genuinely part of the conversation but
  may not appear in the thread for a second or so. A caller that replies and immediately reads the
  thread back to confirm should treat a missing item as "not indexed yet" rather than "lost".

- **`mail.list` and `mail.search` can be paged past the first result page** (#43). Both accept a
  `cursor` and return `nextCursor`, `hasMore`, `sortedBy` and `sortDirection`.

  Previously a response could only report `truncated: true`. That tells a caller there is more mail
  but gives them no way to reach it: re-issuing the same call returns the same first page forever,
  and there was no offset to advance. The honest reading of a truncated response was "some matches
  exist and you cannot see the rest".

  The cursor is a keyset token, not an offset. It records the received time of the last item
  examined, so a page boundary is a position in the ordering rather than a position in a list, and
  mail arriving mid-walk does not shift every subsequent page by one - which is how an offset
  produces a duplicate and a silent miss at the same time. Because received times are not unique, it
  also carries the entry ids seen at exactly that instant and re-scans that band, so tied timestamps
  neither repeat nor disappear.

  A cursor is bound to the query that minted it and is rejected if replayed against a different
  folder, query or filter. `maxCount` is deliberately excluded from that binding, so page size may
  change part-way through a walk. An unreadable or stale cursor is a clean failure rather than a
  silent restart: quietly returning page one would make a caller looping on `hasMore` never
  terminate.

  This is a live keyset walk, not a snapshot. Within the range walked nothing is returned twice and
  nothing is silently skipped; mail deleted mid-walk will not appear, and mail arriving above the
  boundary is not retro-fitted into a page already passed.

### Fixed

- **Meeting invitations are no longer invisible in mail listings** (#32). `mail.list` and
  `mail.search` cast every folder item with `as MailItem` and skipped anything that came back null.
  A meeting request is a `MeetingItem`, an unrelated COM type, so every invitation, cancellation and
  response was silently absent - and the response said nothing about it. A user asking what was in a
  folder got some of it, confidently, with no sign anything had been withheld. In one folder on the
  test mailbox that was seven items out of forty.

  Scheduling items are now listed, and every entry carries `itemType` (`mail`, `meetingRequest`,
  `meetingCancellation`, `meetingResponse` or `other`) because a caller must be able to tell them
  apart: replying to an invitation is not accepting it. A meeting's attendees are rendered into `to`,
  since `MeetingItem` has no `to` of its own and a meeting listed as addressed to nobody is
  misleading.

  Anything still not summarisable is counted in `skippedItemCount` rather than dropped. A listing
  whose numbers do not add up is how "here is what is in your folder" quietly becomes false.

  Responding to invitations, creating meetings with attendees and FreeBusy lookup remain open on #32.

- **The tool no longer claims Outlook is not running while Outlook is running** (#90). Availability
  was determined solely by looking Outlook up in the COM Running Object Table. Outlook does not
  reliably register itself there: with Outlook open, the mailbox loaded and the window in front of
  you, `GetActiveObject` returns `MK_E_UNAVAILABLE` while `CoCreateInstance` on the same ProgID, in
  the same process and as the same user, hands back a working `Application` immediately.

  So every call returned "Outlook does not appear to be running" - a confident, checkable, wrong
  statement about the state of the machine. It also meant every integration test in this repository
  skipped itself, and the suite reported green while verifying nothing against a mailbox.

  Resolution now falls back to attaching via `CoCreateInstance`, which returns the already-running
  instance because Outlook is a single-instance COM server. The fallback is gated on an `OUTLOOK.EXE`
  process actually existing **in the current Windows session**, because `CoCreateInstance` would
  otherwise *launch* Outlook, which this tool must never do. The session filter is load-bearing: a
  COM attach cannot reach another session's Outlook, so counting one would defeat the guard.

- **`folder.list-items` no longer returns an arbitrary subset of a large folder** (#91). It walked
  the folder in store order and stopped at `maxCount`, so 25 items out of 119 were an arbitrary 25 -
  a caller looking for a message they had just created could be told, in effect, that it did not
  exist. Items are now ordered newest-first before truncation, and the response says which property
  the ordering used (`sortedBy`, `sortDirection`) and whether anything was cut (`truncated`). Where
  no orderable property exists the order is reported as unknown rather than presented as an ordering.

- **Replying to or forwarding an unsent draft now explains why it cannot work** (#92). Outlook
  rejects the operation - there is nobody to reply to - but reports it as "Could not send the
  message", which describes an action the caller never requested and gives an agent nothing to act
  on except a retry that cannot succeed. The message now names the cause and points at the draft
  editing actions instead.

- **`mail.search` no longer silently misses terms buried in long message bodies** (#42). The
  free-text `query` matcher fetched the whole body through COM and then discarded everything past
  1200 characters before searching it. A term further in was invisible, and the caller was not told
  the body had been truncated - they were told there was no such mail, which is the one answer a
  search must never give wrongly.

  The truncation also bought nothing. The expensive part is the `mail.Body` COM call, and by the
  time the string was cut that had already happened, so matching the full body costs the same. The
  1200 character limit was pure downside.

  Body matching is now exhaustive over the whole message. Note the trade-off this makes explicit
  rather than hidden: reaching the body means opening every candidate item, so `query` is slow in a
  large folder. Pair it with the structured filters added in #27 - those run inside Outlook and
  decide how many items have to be opened at all. Indexed body search via `AdvancedSearch` remains
  open under #42.

### Added

- **`mail.list` and `mail.search` accept structured filters that Outlook evaluates server-side**
  (#27). Five new parameters - `fromAddress`, `subjectContains`, `receivedAfter`, `receivedBefore`
  and `hasAttachment` - are compiled into a DASL `@SQL=` string and pushed down through
  `Items.Restrict`, joining the `unreadOnly` filter that already worked this way. Previously the only
  way to find mail by sender or date was to ask for a large listing and filter it client-side, which
  could not see past the scan cap: in a busy folder a matching message simply never got read, and the
  caller was told there was none.

  The filter is deliberately built to be over-inclusive rather than exact. `Restrict` runs inside
  Outlook, so anything it wrongly excludes is unrecoverable and surfaces as a confident "no such mail
  exists"; anything it wrongly includes is removed by the client-side check that still runs
  afterwards. Consequently a value containing a DASL `LIKE` wildcard (`%` or `_`) drops that one
  predicate instead of emitting a filter that would distort the match, and date bounds carry a minute
  of slack. Embedded single quotes are escaped by doubling. The `Restrict` call falls back to an
  unfiltered scan if Outlook rejects the filter, and both paths return identical results.

  Verified against classic Outlook 16.0.0.20430 on a live mailbox, including apostrophes, wildcards,
  filter-injection attempts, and AND-combination. `search`'s free-text `query` is unchanged and still
  applied client-side, because `Restrict` cannot see message bodies; server-side full-text search is
  tracked separately as #42.

### Fixed

- **The CLI no longer silently ignores unknown options** (#81). `Spectre.Console.Cli` defaults to
  non-strict parsing, which collects unrecognised options into the remaining-arguments bag instead of
  rejecting them, and `Program.cs` never opted out. A typo therefore produced a confidently wrong
  answer rather than an error: `outlookcli mail list --folder inbox --limit 3` (the real option is
  `--max-count`) returned 25 messages with `"success": true` and exit code 0. That is a direct
  violation of the project's own rule that `success` must match reality, and it is worst for the
  primary consumer - an LLM guessing a plausible flag name is given no signal that it guessed wrong.
  Strict parsing is now enabled, so unknown options fail the command. Parse and runtime failures also
  report as `Command error:` with a `--help` hint instead of `Unhandled error:`, which read like a
  crash for what is usually a typo. Found by running the CLI against a live classic Outlook mailbox.
- **`service status` contract test no longer asserts a field that does not exist.**
  `ServiceRun_ReportsZeroSessionsInitially` asserted a `sessionCount` property inherited from the
  PowerPoint origin, where one session existed per open presentation. The Outlook daemon holds a
  single shared `Application` and has no session concept, so the assertion threw
  `KeyNotFoundException` on every run. Replaced with a test that pins the payload's real shape and
  guards against the stale field returning.
- **The `((dynamic))` cast pre-commit gate was silently not running** (#82). Three defects combined
  so that `pre-commit.ps1` printed `All pre-commit checks passed!` without the check having verified
  anything. First, `scripts/check-dynamic-casts.ps1` is UTF-8 without a BOM and contained em dashes;
  Windows PowerShell 5.1 decodes a BOM-less file using the ANSI code page, mangling those bytes so
  the file no longer parses - and `powershell.exe` is what a git hook invokes by default, so the gate
  had never run for anyone using it. Second, `pre-commit.ps1` caught the resulting error, printed
  `Continuing...`, and still reported overall success; the COM leak check had the same swallow. Both
  now exit non-zero, because a check that cannot run has not passed. Third, the scan used
  `Get-Content` unwrapped, which returns a bare string for a single-line file, and indexing a string
  yields characters - so a single-line `.cs` file was never scanned. All three are fixed and the gate
  is verified to detect an undocumented cast and to accept a documented one under both shells.
- **A timed-out Outlook operation no longer runs after its caller gave up** (#19). Work items were
  queued onto the shared STA dispatcher and executed unconditionally when the thread became free, even
  if the caller had already thrown `TimeoutException` minutes earlier. Outlook operations are not all
  read-only, so an abandoned `mail.send`, `mail.delete` or `folder.create` performed a real mailbox
  side effect that the caller was told had timed out and may legitimately have retried. Work items now
  check their caller's deadline immediately before running and are dropped if it has passed. This
  narrows the window rather than closing it - a caller can still time out in the instant between the
  check and the call - but that residual race is unavoidable without cancellable COM and is orders of
  magnitude smaller than the queue wait it replaces.
- **A wedged dispatcher is now quarantined instead of silently swallowing every later call** (#19). A
  blocking cross-apartment COM call cannot be interrupted: `Thread.Abort` is unsupported on .NET 5+,
  and `OleMessageFilter` only governs *incoming* calls, so it cannot rescue an outgoing call parked on
  a modal Object Model Guard prompt. Previously every subsequent caller queued behind such a call and
  burned its own full `DefaultOperationTimeout` - five minutes each - before reporting a generic
  timeout that named the wrong operation. Once the in-flight operation is past its own deadline, new
  callers are now rejected immediately with a message naming the stuck operation and pointing at the
  likely dialog. The condition clears itself as soon as the blocked call returns.

### Added

- **Issue forms replace the legacy markdown issue templates** (#1). The three templates were markdown,
  so every field in them was a suggestion; a reporter could delete the whole body and submit. They are
  now YAML forms with the fields that actually make an Outlook bug actionable marked required:
  `application.get-status` output, whether Outlook is classic or new, reproduction steps, and expected
  versus actual behaviour. Each form also carries a required acknowledgement not to paste mailbox
  contents, since Outlook issues attract real subjects and addresses. Blank issues are disabled and
  `config.yml` routes security reports to the policy rather than a public issue.
- **Dependabot now covers npm** (#1). `vscode-extension/` has a `package-lock.json` and had no
  Dependabot `updates` entry, so its vulnerability alerts had no update PR to fix them - which is why a
  high-severity advisory sat open with no action. npm is configured for that directory with the same
  weekly schedule, labels and grouping as the existing ecosystems.

  Note that the labels these files reference (`mcp-server`, `npm`, `nuget`, `github-actions`,
  `automated`) did not exist on the repository and have been created. GitHub does not create a missing
  label on demand: an issue form referencing one fails to apply it, and Dependabot refuses the update
  config outright, so the existing `nuget` and `github-actions` entries had been mislabelled all along.

### Changed

- **Removed the `docker` Dependabot ecosystem** (#1). The repository contains no Dockerfile or compose
  file, and the entry was commented ".NET SDK (via global.json if present)", which is not what a docker
  ecosystem watches. It was configuration that could never produce a PR.
- **`docs/CONTRIBUTING.md` documents dependency-update ownership** (#1), points bug and feature
  reporters at the new forms, and warns against hand-bumping `package-lock.json` behind a corporate
  registry mirror, which rewrites `resolved` URLs and can downgrade `integrity` from sha512 to sha1.
- **`Items` collections are now early-bound instead of `object` + `((dynamic))`** (#74). Six locals in
  `CalendarCommands`, `FolderCommands` and `MailCommands` were declared `object?` and then late-bound
  on every use, which cost a DLR call site per `.Count`, per indexer read and per `Sort`/`Restrict`,
  and gave up compile-time checking on `Outlook.Items` members that have always been in the PIA.
  They are `Outlook.Items?` now, and 15 of the 16 `((dynamic))` casts are simply gone. The one
  genuinely late-bound site remains: `CreateFolderItemInfo` reads `MessageClass`, `Subject`,
  `FullName` and `Name` off an item whose type it could not identify, and the PIA gives those
  classes no common interface. It is now a single named `dynamic` local carrying the explanation.

### Fixed

- **`scripts/check-dynamic-casts.ps1` no longer fails on `master`** (#74). It reported 16 undocumented
  casts on every run, so the pre-commit hook could not pass on an unmodified checkout. A gate that is
  always red trains people to bypass the hook, which is how a real defect gets through. All ten
  pre-commit checks now pass end to end.

### Removed

- **Last PowerPoint residue in build and repo config** (#70 follow-up). `dotnet build` no longer
  makes the `Microsoft.Office.Interop.PowerPoint` PIA available at all: its `PackageVersion` entry
  is gone from `Directory.Packages.props` (no project had referenced it since #73). Also deleted
  `scripts/build_complex_test.py`, a dead harness that drove the MCP server against a `.pptx` on
  the desktop through a hard-coded path to a checkout that no longer exists, plus the `*.pptx` /
  `*.pptm` rules in `.gitattributes` and the stale Excel/PowerPoint test-asset exceptions in
  `.gitignore` that pointed at directories deleted in #73.

### Fixed

- **`scripts/Stop-OutlookMcpProcesses.ps1` never found the CLI.** It probed
  `bin\{Debug,Release}\net10.0-windows\outlookcli.exe`, but the projects target `net9.0-windows`,
  so the pre-build cleanup always fell through to the named-pipe fallback instead of asking the
  service to shut down cleanly.

- **The inherited PowerPoint session/batch layer is gone** (#12, #65, #70). The product owner
  directed that nothing PowerPoint remain, which overrules ADR-002's earlier decision to retain the
  layer as dormant infrastructure. ADR-002 has been amended to record the reversal and its reasons.
  Deleted:
  - `src/OutlookMcp.ComInterop/Session/`: `PptSession`, `PptBatch`, `PptContext`, `IPptBatch`,
    `SessionManager`, `PptShutdownService`, `ResiliencePipelines`, and the 8 integration test files
    that went with them. Those tests launched PowerPoint on any unfiltered `dotnet test` (#70).
  - The CLI `session` command branch (`create`, `open`, `close`, `list`, `save`) and
    `SessionCommands.cs`. No Outlook operation ever used it.
  - The service's session RPC handlers, `SessionManager` property, `SessionCount` field on
    `ServiceStatus`, and the session UI in the CLI daemon's tray icon.
  - The `[NoSession]` attribute and every session branch in the four source generators. All five
    Outlook service interfaces carried `[NoSession]`, so the generated `session_id` MCP parameter,
    the `RequiresSession` CLI plumbing, and the `IPptBatch` dispatch argument were dead code for the
    shipping product. **The generated tool and CLI surface is unchanged.**
  - `scripts/refactor-service.ps1`, an unreferenced one-off migration script.
- **`Microsoft.Office.Interop.PowerPoint` dropped entirely**: removed from `OutlookMcp.Core` (after
  #26 nothing in Core referenced a PowerPoint type) and from `OutlookMcp.ComInterop` (its only
  consumer was the session layer). Also removed from `dependency-review.yml`'s allow-list.
  `Microsoft.Office.Interop.Outlook` remains.
- **`llm-tests/` deleted** (#68). Every scenario in the pytest/`pytest-aitest` harness targeted the
  presentation surface removed by #26 - charts, tables, ranges, slides, styling - so the suite
  tested nothing that exists. `scripts/Test-LlmRegressionGate.ps1`, the `run_llm_gate`
  `workflow_dispatch` input in `integration-tests.yml`, and
  `.github/instructions/llm-testing-philosophy.instructions.md` went with it. If LLM-behaviour
  testing is wanted again it should be designed against the five Outlook tools from scratch.
- **`FileAccessValidator` deleted**: OLE2/IRM-container detection for Office *documents*, with no
  caller anywhere in the Outlook surface.
- **PowerPoint file-format constants and quit timeouts removed** from `ComInteropConstants`, which
  now carries only the three values Outlook actually uses.
- **The `RunType=OnDemand` CI step and its `workflow_dispatch` input** were removed from
  `integration-tests.yml`: no test carries that trait any more.

### Changed

- **Naming debt #12 closed.** The last `Ppt*` identifiers are gone: the generated MCP tool types are
  now `Outlook{Category}Tool` (the wire-level tool names - `mail`, `calendar`, `folder`,
  `attachment`, `application` - are unchanged, so **no MCP client sees a difference**), the
  hand-written base class is `OutlookToolsBase`, the daemon RPC contract is `IOutlookDaemonRpc`, and
  the generated prompt class is `OutlookSkillPrompts`.
- **`ServiceBridge.SendAsync` and the generated `RouteAction`/`Forward*` methods lost their
  `sessionId` parameter.** It was threaded through the whole MCP call path and was always the empty
  string. Internal only; the tool schemas are unchanged.

### Fixed

- **The pre-commit hook was unusable for everyone who installed it** (#65): check 6 ran
  `scripts/Test-CliWorkflow.ps1`, which drove the PowerPoint CLI surface deleted by #26 - creating a
  `.pptx`, adding slides and shapes - so it failed immediately, and it launched PowerPoint while
  doing so. Rewritten against the real Outlook CLI surface. The new script asserts only things that
  hold whether or not Outlook is running: `diag ping` answers, `diag echo` round-trips a parameter
  through the pipe, `diag outlook` and `service status` return well-formed JSON, an unknown action
  exits non-zero, `--output` writes no file on failure, and - the important one - `application
  get-status` reaches the generated dispatch surface and its process exit code **agrees** with the
  `success` field in its payload, which is a standing regression guard for #63. It never launches an
  Office application.
- **`OleMessageFilterTests.MessagePending_ReturnValue_MustBe_WaitDefProcess` failed on `master`** (#59):
  - ROOT CAUSE: the **test**, not the implementation. It declared
    `PENDINGMSG_WAITDEFPROCESS = 1` and `PENDINGMSG_WAITNOPROCESS = 2`. The Win32 `PENDINGMSG`
    enumeration in `objidl.h` defines the opposite: `CANCELCALL = 0`, `WAITNOPROCESS = 1`,
    `WAITDEFPROCESS = 2`. `OleMessageFilter.MessagePending` correctly returns 2, which is
    `WAITDEFPROCESS` - exactly the value the test's own comment said it wanted - and the test then
    asserted that 2 was the forbidden value.
  - FIX: corrected the two constants and the surrounding rationale, and cited the Win32 enumeration
    in the test so the values cannot silently drift again. The filter implementation is unchanged.
- **Pre-commit branch guard never fired**: `scripts/pre-commit.ps1` blocked commits to a branch
  named `main`, but this repository's default branch is `master`, so the Rule 6 guard silently
  passed on every direct commit to the default branch.
- **CLI returned exit code 0 when an operation failed** (#63): `outlookcli` printed
  `{"success": false, "errorMessage": "..."}` on stdout and then exited 0, so every script, CI step,
  and agent that branched on `$LASTEXITCODE` treated a failed Outlook operation as a success.
  - ROOT CAUSE: `ServiceCommandBase.ExecuteAsync` checked `response.Success`, which is a
    *transport*-level flag meaning "the daemon replied and routed the request". The operation's own
    outcome is carried in the `success` property of the JSON payload inside `response.Result` and was
    never inspected. Argument-parse and validation errors already returned 1, which is exactly why
    the gap went unnoticed.
  - FIX: added `ServiceCommandBase.ResolveExitCode`, which inspects the result payload and returns 1
    only on an explicit `success: false`. Payloads with no `success` property (bare arrays, ordinary
    read results), empty payloads, and non-JSON output are still treated as success, because the
    daemon has already confirmed the call ran and guessing there would turn valid results into
    spurious errors.
  - Also fixes a Rule 0 violation on the `--output` path: `WriteOutputToFile` wrote the failure
    payload into the target file and then announced `{"success": true, "outputPath": ...}` on stdout.
    A failed operation now surfaces the error and writes no file.
- **`ListPrompts_ReturnsOnlyOutlookPrompts` asserted a prompt name that no longer exists**: #34
  replaced `skills/shared/outlook_agent_mode.md` with `outlook-workflows.md`, which changes the
  generated prompt name from `outlook_agent_mode_guide` to `outlook_workflows_guide`. The test is an
  integration test, so no CI path currently runs it and the break was silent.

### Changed

- **`.github/instructions/` rewritten for Outlook COM** (#62): the instruction files are
  auto-loaded into every Copilot session in this repository, so they were actively steering
  contributors toward a PowerPoint batch/session API that no Outlook code path uses.
  - `ppt-com-interop.instructions.md` and `ppt-com-patterns-guide.instructions.md` were replaced
    by a single `outlook-com-interop.instructions.md` documenting the real execution model:
    `OutlookInteropRunner.Execute` on the shared `OutlookDispatcher` STA thread, strongly-typed
    `Microsoft.Office.Interop.Outlook` rather than `dynamic`, try-finally release, the
    never-final-release-the-shared-Application rule (#19), and the Object Model Guard posture.
  - `copilot-instructions.md`, `critical-rules.instructions.md` (Rules 1b, 5, 7, 9, 12, 14, 16,
    18, 22, 24, 28, 30), `architecture-patterns`, `testing-strategy`, `mcp-server-guide`,
    `coverage-prevention-strategy`, `mcp-llm-guidance`, `development-workflow` and `meta` were
    retargeted to Outlook.
  - Corrected claims that became false after #5/#11 and #26: `ToolActions.cs` and
    `ActionExtensions.cs` no longer exist (the MCP tool, its action enum and the CLI command are
    all generated from `[ServiceCategory]`/`[ServiceAction]`), the test `Feature` traits no longer
    include `Slide`/`Shape`/`Text`/`VBA`/`Screenshot`, and the surface is five tools, not 19.
  - Added a warning that stale generator output under `obj/` is not the source of truth.
- **Orphaned `.github/agents/copilot-instructions.md` deleted**: an auto-generated spec-kit stub
  whose generator (`.specify/`) is not in the repository. Nothing referenced it and its content
  was a placeholder plus stale "PowerPoint COM automation via `dynamic`" technology notes.
- **Rule 30 / ADR-001 narrowed: the ban is on mocked-COM unit tests, not on all unit tests** (#37):
  the rule said "NEVER write unit tests" while 16 files sat under `tests/**/Unit/`. A rule that the
  codebase visibly contradicts gets ignored wholesale rather than obeyed selectively, so #37 asked
  for an explicit decision. Chose to narrow the rule rather than delete all 16 or retire it. A unit
  test is now permitted only if it touches no COM object at all (not a real one, not a `null!`
  stand-in, not a mock), its subject is genuinely pure, it would fail if the logic were wrong rather
  than only if .NET were broken, and it is traited `Category=Unit` under `tests/**/Unit/`. ADR-001
  now carries the normative list of permitted files; adding to it requires amending the ADR in the
  same PR. Stated plainly that a permitted unit test never counts as coverage for a COM operation.

### Removed

- **Dead PowerPoint harnesses deleted** (#61): removed `eval/` (82 files) and
  `src/OutlookMcp.Agent/` (12 files). Both were PowerPoint deck-generation tooling inherited from
  the pre-rename repository; neither is in `OutlookMcp.sln`, neither ships in any package, and
  nothing in the Outlook product referenced them. Their documentation
  (`docs/AGENT-CLIENT.md`, `docs/ARCHETYPE-PIPELINE.md`), the `scripts/Sync-EvalAssets.ps1` helper,
  and the orphaned `tests/OutlookMcp.Core.Tests/TestAssets/ReferenceCatalog/` fixture set (whose
  only consumers, `DesignReferenceCatalogTests` and `ReferenceCatalogFixture`, went in #26) are
  removed with them.
  - Dropped the `eval-tools` job and the `eval/` path triggers from `.github/workflows/node-ci.yml`.
  - Dropped the nine `pkg:npm/%40github/copilot*` entries from `dependency-review.yml`'s license
    allow-list. Those existed solely because `@github/copilot-sdk` ships under GitHub's proprietary
    terms, which no SPDX allow-list can express; with both consumers gone, the repository no longer
    depends on it and the exception is no longer needed.
- **Three mocked-COM unit tests deleted** (#37): `ComUtilitiesTests`, `ComUtilitiesExtendedTests`,
  and `PptContextTests`. All three claimed to exercise COM behaviour while passing `null!` or plain
  strings where a COM object belonged; one test was named `Release_WithComObject_DoesNotThrow`
  directly above a comment conceding that no COM object was involved. `OleMessageFilterTests` also
  fails the new test, but is deliberately left in place because it is currently failing on `master`
  (#59) and deleting a failing test is how real defects get lost.

- **Dead documentation deleted** (#34): `.github/ISSUE_TEMPLATE/breaking-changes-issue.md` (it had no
  YAML front matter, so it was never a real issue template, and every document and tool it referenced
  had already been deleted) and `tests/OutlookMcp.Core.Tests/docs/DATA-MODEL-SETUP.md` (documented
  Data Model tests that no longer exist and that nothing references).
- **Legacy PowerPoint command surface deleted** (#26): removed all 33 inherited PowerPoint command
  domains from `src/OutlookMcp.Core/Commands/` (69 files: `Accessibility`, `Animation`, `Background`,
  `Chart`, `Comment`, `CustomShow`, `Design`, `DocumentProperty`, `Export`, `File`, `HeaderFooter`,
  `Hyperlink`, `Image`, `Master`, `Media`, `Notes`, `PageSetup`, `Placeholder`, `PrintOptions`,
  `Proofing`, `Section`, `Shape`, `ShapeAlign`, `Slide`, `SlideImport`, `Slideshow`, `SlideTable`,
  `SmartArt`, `Tag`, `Text`, `Transition`, `Vba`, `Window`), leaving only the five Outlook domains
  (`Application`, `Attachment`, `Calendar`, `Folder`, `Mail`) plus the `OutlookInterop` helper.
  - **This also fixes the CLI/MCP asymmetry left by #23.** #23 removed the legacy tools from the MCP
    server only, via an explicit allow-list in `Program.cs`; the CLI still registered all 33 legacy
    domains because `CliCommandRegistration` is source-generated from *every* `[ServiceCategory]`
    interface in Core. Deleting the interfaces removes the generated CLI commands automatically, so
    the CLI project needed no edits. The generated CLI surface is now exactly the five Outlook
    categories.
  - Also removed the now-dead `src/OutlookMcp.Core/Data/` design-catalog provider and its three
    embedded JSON/Markdown resources (unreferenced once the `Design` domain went), and the legacy
    MCP tool files `PptFileTool.cs`, `PptTools.cs`, and `PptResourceProvider.cs`.
  - `PptToolsBase.cs` is **kept**: the generated Outlook tools call into it. Its `Ppt*` name is
    part of the #12 naming debt and is out of scope here.
  - Per the amended ADR-002 (below), `ComInterop/Session/*` is **retained** as dormant infrastructure.
- **Legacy tests removed**: `DesignReferenceCatalogTests`, `ReferenceCatalogFixture`,
  `ShapeHelpersTests`, `PptDesignToolTests`, `PptFileToolTests`, `DesignCommandTests`, and
  `ParameterValidationTests` (which covered only deleted PowerPoint domains).

### Changed

- **Documentation and packaging purged of PowerPoint** (#34): the markdown surface went from 90 files
  with 1041 PowerPoint references down to 44 files, and every survivor is now deliberate. Rewrote
  `README.md`, `FEATURES.md`, `SECURITY.md`, the `docs/` set, `tests/README.md`, every project
  `README.md`, the `mcpb/` and `vscode-extension/` packaging metadata, the `packages/*` skill
  READMEs, `examples/README.md`, and all four GitHub issue/PR templates.
  - **The two shipped MCP prompts were the most urgent part.** Every `.md` under `skills/shared/` is
    code-generated into an `[McpServerPrompt]` and copied into both skill reference folders.
    `behavioral-rules.md` and `outlook_agent_mode.md` were entirely PowerPoint, so the server was
    actively shipping LLM instructions for `slide()`, `shape()`, and `window()` tools that #26
    deleted. `behavioral-rules.md` was replaced with nine Outlook rules, and `outlook_agent_mode.md`
    was replaced by a new `outlook-workflows.md` (its premise -- window show/hide -- has no Outlook
    equivalent). Slide-design prompts moved to `eval/skills/`, out of the shipped prompt set.
  - **Breaking for anyone activating the integration runner:** the repository variable
    `ENABLE_POWERPOINT_INTEGRATION_CI` is renamed to `ENABLE_OUTLOOK_INTEGRATION_CI`, and the
    self-hosted runner label `powerpoint` becomes `outlook`. `docs/AZURE_SELFHOSTED_RUNNER_SETUP.md`
    is updated to match and now also requires a working Outlook mail profile on the runner host.
  - `skills/outlook-cli/references/cli-commands.md` was a stale stub listing 24 deleted domains with
    empty sections; it is regenerated from the real `outlookcli --help`.
  - `src/OutlookMcp.ComInterop/README.md` keeps its PowerPoint wording **on purpose** and now states
    so explicitly, splitting the library into the active Outlook path (`OutlookDispatcher`,
    `ComUtilities`, `OleMessageFilter`) and the dormant retained `Ppt*` session layer.
  - Also fixed PowerPoint wording in shipped CLI help text (`Program.cs`, `CliServiceTray.cs`,
    `SessionCommands.cs`, `ListActionsCommand.cs`), so this is not a docs-only change.
  - Deferred by design: `.github/instructions/` and `.github/copilot-instructions.md` (#62), and
    `eval/` plus `src/OutlookMcp.Agent/` (#61). `CHANGELOG.md` and the ADRs keep their PowerPoint
    references as historical record.

- **`vscode-extension-ci.yml` renamed to `node-ci.yml`**: the workflow was renamed to `Node Projects
  CI` when it took on `eval/` tooling, but its filename still claimed to be extension-only.

- **`CoreCommandsCoverageTests` and `ActionValidatorTests` are now reflection-driven**: both
  previously hard-coded the list of command domains, and `CoreCommandsCoverageTests` had silently
  drifted -- it omitted `ICalendarCommands` entirely, along with the `AttachmentAction` and
  `CalendarAction` mapping tests. They now discover categories from the `[ServiceCategory]`
  attribute and the generated `_CliCategoryMetadata`, so adding or removing a domain cannot silently
  drop coverage. Both include a guard assertion so the theories cannot pass vacuously.
- **CLI and MCP help text no longer advertises deleted commands**: `outlookcli actions` described a
  session workflow ending in `slide list --session abc`, a command that no longer exists. It now
  describes the Outlook commands and notes that the session commands drive the retained
  presentation-session layer and are not required by any Outlook command. The MCP `Program` summary
  no longer claims to host "inherited legacy presentation tools".

- **ADR-002 amended: `ComInterop/Session/*` is retained, not deleted** (#40, #26): the ADR previously
  stated that `ComInterop/Session/*` (`PptSession`, `PptBatch`, `PptContext`, `SessionManager`,
  `ResiliencePipelines`, `PptShutdownService`) "should be deleted wholesale" once #26 removes the
  legacy PowerPoint command surface. That directly contradicted #26's own acceptance criteria, which
  require the layer to be **retained**. The conflict is resolved in favour of #26: the layer is kept
  as dormant infrastructure, because it is the only mature COM plumbing in the repository (dedicated
  STA pump, Polly resilience pipelines, operation tracking, timeout handling, named-pipe daemon) and
  is covered by 8 integration test files, so its reconstruction cost exceeds the cost of carrying it.
  Rewrote the "Fate of `ComInterop/Session/*`" and "Naming Plan" sections, added the rejected
  wholesale-deletion option to "Alternatives Considered", and restated the "generalize this layer"
  rejection, whose original rationale partly rested on the now-false premise that the layer was
  scheduled for deletion. **The ADR's central decision is unchanged**: Outlook executes via
  `OutlookDispatcher` (#20), never via `PptSession`/`PptBatch` -- retention is not reuse. Also noted
  that the retained `Ppt*` types should *not* be renamed to `Outlook*` under #12, since Outlook does
  not use them and an `Outlook*` prefix would misdescribe them; a product-neutral rename is deferred
  to the re-scoping of #12.

### Added

- **`Node Projects CI` workflow**: runs `npm ci` and `tsc` for `vscode-extension/`, and `npm ci` for `eval/`, on pull requests. Previously nothing installed those dependencies on a PR -- only the manually dispatched `release.yml` did -- so an uninstallable lockfile would first have surfaced mid-release.

### Changed

- **BREAKING — Project and package rename `PptMcp.*` → `OutlookMcp.*`** (#5): every project under `src/` and `tests/`, the solution file, all assembly names, root namespaces, `using` directives, NuGet package IDs, build scripts, and CI workflows were renamed so this repository publishes only under its own identity.
  - NuGet package IDs: `PptMcp.McpServer` → `OutlookMcp.McpServer`, `PptMcp.CLI` → `OutlookMcp.CLI`, `PptMcp.Core` → `OutlookMcp.Core`, `PptMcp.ComInterop` → `OutlookMcp.ComInterop`
  - dotnet tool commands: `pptcli` → `outlookcli`, `mcp-ppt` → `mcp-outlook`
  - npm packages: `ppt-mcp-skill` → `outlook-mcp-skill`, `ppt-cli-skill` → `outlook-cli-skill`
  - Release artifacts: `PptMcp-<version>.vsix` → `OutlookMcp-<version>.vsix`, `ppt-skills-v<version>.zip` → `outlook-skills-v<version>.zip`
  - Skill generation now writes to `skills/outlook-mcp` and `skills/outlook-cli`; the stale `skills/ppt-mcp` and `skills/ppt-cli` directories were removed
  - Repository self-references updated from `trsdn/mcp-server-ppt` to `trsdn/mcp-server-outlook`

### Added

- **Ported `folder.resolve-path`, `folder.list-items`, `mail.set-subject`, `mail.set-body`, `mail.set-recipients` from the orphaned `feature/outlook-parity-slices` branch** (#28): that branch was cut before the `PptMcp` → `OutlookMcp` rename (#5) landed on `master` and was never merged, leaving genuinely useful capabilities stranded. Rather than merging or rebasing the whole branch — which also carries untested Contacts CRUD and Application explorer/inspector context work — only the two smallest, self-contained, highest-value slices were cherry-picked and re-ported onto current `OutlookMcp.*` namespaces: `folder.resolve-path`/`folder.list-items` (new mailbox folder navigation, not previously exposed) and `mail.set-subject`/`mail.set-body`/`mail.set-recipients` (draft editing, a prerequisite called out by #28 for future draft-editing (#15) and meeting-creation (#14) work since `Recipients` mutation was otherwise unavailable). The Contacts and Application explorer/inspector slices were deliberately **not** ported in this pass — they need new result types and additional test coverage and are tracked separately. Added `[SkippableFact]` smoke tests for all five new actions in `OutlookSeedSmokeTests`, following the existing integration-test-only pattern (Rule 30).

- **Detect and report classic-vs-new Outlook for Windows** (#35): the new Outlook for Windows (the modern packaged Mail & Calendar replacement) has no COM object model and cannot be automated by this server, but previously a missing `Outlook.Application` ProgID always produced a generic "Outlook not installed" message, even when new Outlook was present and running. Added `OutlookInstallationDetector` (a pure, zero-COM-dependency utility using `Type.GetTypeFromProgID` and new-Outlook package registry checks) that distinguishes `NotInstalled`, `ClassicDesktop`, `NewOutlookOnly`, and `Unknown`, plus a same-process-integrity-level check via `WindowsPrincipal`. `application.get-status` now reports `outlookFlavor`/`processElevated`, and the not-installed error path in `OutlookInteropRunner` surfaces a flavor-specific, actionable message (e.g. telling the user new Outlook is installed but unsupported, vs. classic Outlook installed but not running). Added `outlookcli diag outlook` to surface the same detection standalone. Updated `README.md`'s Requirements section to explicitly call out that classic Outlook desktop (not new Outlook for Windows) is required.

- **`tools/list`/`prompts/list` no longer expose the 33 inherited legacy PowerPoint tools, the hand-written `file` tool, or the 3 PowerPoint-only skill prompts** (#23, part of #11): `Program.cs` previously registered every `[McpServerToolType]`/`[McpServerPromptType]` in the assembly via `.WithToolsFromAssembly()`/`.WithPromptsFromAssembly()`, so all 38 generated tool types (only 5 -- `application`, `attachment`, `calendar`, `folder`, `mail` -- are Outlook) and all 5 skill prompts (only 2 -- `behavioral_rules_guide`, `outlook_agent_mode_guide` -- are Outlook) were offered to every MCP client. This flooded `tools/list`/`prompts/list`, wasted LLM context, and invited mis-selection; the hand-written `file` tool (`Destructive = true`, no session requirement) was also a live, unintended gateway into `PptBatch` → `PowerPoint.Application`. Replaced assembly-wide discovery with an explicit allow-list in `Program.cs` using the reflection-based `WithTools(IEnumerable<Type>)`/`WithPrompts(IEnumerable<Type>)` overloads (the generic `WithTools<T>`/`WithPrompts<T>` require non-static types, but the generated tool/prompt classes are static). Narrowed the `PptSkillPrompts.g.cs` MSBuild generator target and its embedded-resource `ItemGroup` in `OutlookMcp.McpServer.csproj` to only `behavioral-rules.md` and `outlook_agent_mode.md`, leaving the PowerPoint-only skill docs (`slide-design-*.md`, `generation-pipeline.md`) on disk for the legacy `eval/` harness and skill-based clients' `references/` copies, but out of the MCP prompt surface. Updated `McpServerIntegrationTests`'s `ExpectedToolNames` regression list from the old 38-tool set down to the 5 Outlook tools, mirrored the same allow-list into the test's own DI setup (which previously called `WithToolsFromAssembly` independently of `Program.cs`), replaced its `file`-tool invocation test with an `application get-status` invocation test, and added `ListPrompts_ReturnsOnlyOutlookPrompts` as the regression guard for the prompts half of this fix. Deleted `PptFileToolOperationTrackingTests.cs`, which exclusively exercised the now-unregistered legacy `file` tool's session-tracking behavior through the real MCP server and could only fail once `file` was correctly de-registered. The 33 legacy PowerPoint tool types (and the hand-written `file` tool) still compile and are unaffected otherwise -- deleting that code is a later step of #11.


- **ADR-002: Outlook COM execution model** (#40): recorded the decision to build a purpose-built Outlook STA dispatcher rather than reusing PowerPoint's `PptBatch`/`PptSession` session layer, since Outlook has a single shared always-running `Application` (identified by `entryId`/`storeId`) rather than PowerPoint's per-file open/close model. Documents the fate of `ComInterop/Session/*` (retained for PowerPoint only, deleted alongside the legacy surface in #26) and the naming plan for the new dispatcher.

- **Serialize all Outlook COM access behind a single long-lived STA dispatcher** (#20, P0): previously every `OutlookInteropRunner.Execute` call spawned its own short-lived STA thread, registering/revoking `OleMessageFilter` on every single call — this could allow overlapping COM re-entrancy between operations and paid the thread-spin-up cost repeatedly. Added `OutlookDispatcher` (`OutlookMcp.ComInterop.Session`), a process-wide singleton exposing a generic `Execute<T>(operationName, operation, timeout)` that queues work onto one dedicated STA thread via a **bounded** `Channel<Func<Task>>` (capacity 32, blocking on full — giving explicit back-pressure per the issue's acceptance criteria, unlike PowerPoint's unbounded `PptBatch` channel). `OleMessageFilter.Register()`/`Revoke()` now run exactly once for the dispatcher's whole lifetime instead of per call. `OutlookInteropRunner.Execute` now delegates to `OutlookDispatcher.Shared` instead of owning per-call thread lifecycle; its public signature and all 7 existing call sites (`ApplicationCommands`, `MailCommands`, `CalendarCommands`, `AttachmentCommands`, `FolderCommands`) are unchanged. Added unit tests asserting overlapping concurrent callers are serialized onto a single thread, that a thrown exception propagates without wedging the dispatcher, and that a timed-out operation surfaces `TimeoutException` while leaving the dispatcher usable afterward. Note: as with `PptBatch`, a timeout does not interrupt an in-flight COM call already running on the STA thread — true cancellation/quarantine of stuck operations is tracked separately (#29).

- **`mail.send` confirmation gate and at-most-once retry safety** (#29, P0): sending mail is a destructive, one-way action, but `mail.send` previously fired `MailItem.Send()` unconditionally on any call, and a client-side timeout left the caller unable to tell whether the message actually sent — retrying blindly risked a duplicate send. `MailCommands.Send` now requires an explicit `confirm=true` and is refused with a clear error otherwise. Added an optional `operationId` parameter: results are cached in-process per `operationId` (a `ConcurrentDictionary`, not persisted across restarts), so a caller retrying with the same `operationId` after a timeout replays the first attempt's outcome instead of re-invoking `MailItem.Send()`. Added `MailSendResult.Indeterminate`: if the dispatcher's `TimeoutException` fires while a send may already be in flight, the result now reports `Indeterminate = true` (distinct from `Success = false`) with guidance to re-check via `mail.read` before deciding whether to resend, rather than looking like a definite failure. Added unit tests covering the confirmation gate.

### Fixed

- **CI: `dependency-review` failed on the MS-PL license of the test-only `Xunit.SkippableFact` dependency**: the license gate allow-lists SPDX identifiers plus a small set of package-level exceptions, and `Xunit.SkippableFact` (added in #22 so Outlook integration tests skip loudly instead of passing silently) declares MS-PL. Added `pkg:nuget/Xunit.SkippableFact` to `allow-dependencies-licenses`. MS-PL is OSI-approved and permissive, is already accepted transitively via `StreamJsonRpc`'s `Apache-2.0 AND MIT AND MS-PL` expression, and this is a development-scope dependency that is never redistributed.

- **CI: every workflow ran GitHub Actions that still target the deprecated Node 20 runtime**: GitHub now force-runs `node20` actions on Node 24 and emits a deprecation warning on every job. Bumped all first-party and third-party actions to their first Node 24 major so the workflows run on a supported runtime rather than a forced-compatibility shim: `actions/checkout` v4→v5, `actions/setup-dotnet` v4→v5, `actions/upload-artifact` v4→v6, `actions/download-artifact` v4→v7, `actions/setup-node` v4→v5, `actions/dependency-review-action` v4→v5, and `astral-sh/setup-uv` v4→v7. Deliberately chose the earliest Node 24 major of each action rather than the newest release, to pick up the runtime change without also absorbing unrelated behavioural breaks (notably `download-artifact@v8`, which stops auto-unzipping downloaded artifacts). `actions/stale@v10` and `github/codeql-action@v4` were already on Node 24. `HaaLeo/publish-vscode-extension@v2` remains on Node 20 because upstream has published no Node 24 major; it is release-only, `continue-on-error`, and now carries an inline comment explaining why it cannot be bumped yet.

- **`mail.list`/`mail.search` had a silent, undocumented 500-item scan cap** (#27, P1, partial): `ExecuteMailList` scanned at most `Clamp(maxCount * 10, 25, 500)` items and matched `unreadOnly`/free-text queries entirely client-side; anything beyond that cap was invisible and the tool reported success with an empty or partial list -- an LLM reads that as "no such mail exists", the worst kind of false negative for an agent tool. `unreadOnly` is now pushed down via `Items.Restrict("[Unread] = true")` (DASL) so Outlook's own indexed filter finds unread messages regardless of folder size, with a fallback to the prior client-side check only if `Restrict` throws. The old fixed 500-item cap is replaced with a much larger safety-net-only limit (5000) that exists purely to bound pathological scans, not to define "found". `MailListResult` gained `ScannedCount` and `Truncated`: `Truncated` is explicitly `true` whenever the scan stopped before exhausting the candidate set (whether from `maxCount` or the safety limit), so a short/empty `Messages` list can no longer be silently misread as "nothing matches". **Not yet done** (tracked as remaining #27 scope): `mail.list` does not yet use `Folder.GetTable` for summary projections, `mail.search`'s free-text query does not yet use `Application.AdvancedSearch` for body full-text search, structured DASL predicates (sender/date/category/importance/has-attachment) are not yet exposed as parameters, and there is no paging/continuation token -- `mail.search`'s body-substring matching therefore still relies on client-side scanning bounded by the new safety limit rather than an indexed search.

- **Outlook Object Model Guard denials were unmodelled and silently swallowed** (#30, P0, security): the Outlook Object Model Guard (OMG) raises a modal security prompt for out-of-process, untrusted callers touching protected members (`SenderEmailAddress`, `Recipients`, `MailItem.Send()`, etc.); a denied/unanswered prompt previously looked identical to any other COM failure, and `MailCommands.SafeGet`'s bare `catch { return null; }` made a blocked `SenderEmailAddress` read indistinguishable from "this mail has no sender" — precisely the pattern Rule 22 forbids. Also, `OutlookInteropRunner.GetOrCreateApplication` fell back to `Activator.CreateInstance` when no running Outlook instance could be found via `GetActiveObject`, deliberately creating an untrusted instance that is *more* likely to trigger OMG and conflicts with Outlook's single-instance model. Removed the `Activator.CreateInstance` fallback entirely — a missing running instance now fails with actionable guidance, distinguishing a plain "Outlook not running" from an elevation mismatch (this process running elevated while Outlook runs unelevated, which makes `GetActiveObject` fail with `MK_E_UNAVAILABLE`-shaped errors) via `OutlookInstallationDetector.IsCurrentProcessElevated()`. Added `OutlookInteropRunner.IsObjectModelGuardDenial` to classify `COMException`s consistent with an OMG denial (`E_ABORT`) and surface a distinct, actionable error instead of a generic COM failure. `MailCommands`' security-sensitive property reads (`SenderEmailAddress`, `To`, `Cc`, `Bcc`) now record which specific properties were blocked in a new `AccessDenied` result field on `ActiveMailResult`/`MailSummaryInfo`, rather than returning an ambiguous `null`. Extended `outlookcli diag outlook` with an `objectModelGuard` section documenting the OMG posture and elevation-mismatch risk. Documented the (partial) OMG posture in `docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md`. A full dispatcher-level allowed/denied/cancelled outcome enum threaded through every command is not yet done and remains a candidate follow-up.
- **16/16 Outlook integration tests silently reported `Passed` with zero assertions executed** (#22): `OutlookSeedSmokeTests` guarded every fact with `if (!EnsureOutlookAvailable()) { return; }`, so on any machine or CI runner without a running classic Outlook desktop instance, the suite reported 16/16 green having exercised nothing — violating the repo's "tests must fail loudly, never silent" rule and masking #16's finding that no Outlook CI runner exists. Switched every guarded fact to `[SkippableFact]` (via the `Xunit.SkippableFact` package) and changed `EnsureOutlookAvailable()` to throw `SkipException` instead of returning `false`, so the same runs now correctly report `Skipped`. Added `scripts/check-outlook-tests-not-all-skipped.ps1`, wired into the CI Outlook job, which fails the build if every Outlook test in a TRX result skipped — a 100% skip rate on a runner that is supposed to have Outlook now fails the job instead of manufacturing false confidence.

- **Pre-commit and CI coverage gates were regex-scraping a file that no longer exists, silently reporting false-green** (#25, P1): `audit-core-coverage.ps1` and `check-mcp-core-implementations.ps1` (run in `pre-commit.ps1` and both `build-cli.yml`/`build-mcp-server.yml` CI workflows) read `src/OutlookMcp.Core/Models/Actions/ToolActions.cs`, a hand-authored enum file that predates the move to Roslyn source generators (#5/#11) -- actions are now generated directly from `[ServiceAction]` attributes on Core interfaces, and that file was never re-created. `audit-core-coverage.ps1` degraded to reporting "Total Core Methods: 0, Total Enum Values: 0, Coverage: N/A ... 100% coverage maintained!" on every run — a gate that could never fail was worse than no gate, and it had been silently rubber-stamping every commit and CI build since the migration. `check-mcp-core-implementations.ps1` and two dead CLI-coverage scripts (`check-cli-coverage.ps1`, `check-cli-action-coverage.ps1`, plus the already-orphaned `audit-cli-actions.ps1`) hard-failed on the missing file instead, which is safer but still broken. Deleted all five obsolete scripts; `pre-commit.ps1` and both CI workflows now run the existing reflection-based `CoreCommandsCoverageTests` (which already enumerates the live Outlook Core interfaces/generated enums and genuinely fails on a real gap) instead. Also fixed a real bug in `check-cli-settings-usage.ps1` (now wired into `pre-commit.ps1`, where it previously was not): its `Settings`-class-body regex ran unanchored to end-of-file instead of stopping at the class's closing brace, so it flagged unrelated sibling DTO properties (e.g. `BatchCommand`'s `BatchEntry`/`BatchResult` fields) as "unused Settings properties" -- fixed via brace-balanced extraction. Updated `docs/PRE-COMMIT-SETUP.md`, `.github/copilot-instructions.md`, and `.github/instructions/coverage-prevention-strategy.instructions.md` accordingly. `CoreCommandsCoverageTests`, `ActionEnumCompletenessTests`, and the CLI/MCP Server E2E smoke tests were already re-pointed at the Outlook surface prior to this change and needed no further work.

- **`mail.reply`/`mail.reply-all`/`mail.forward` could not be targeted and were unusable headlessly** (#36, P1): these were the only mail actions with no `entryId`/`storeId`/`useActiveMail` parameters, resolving their source message only via `ActiveInspector`/`Explorer.Selection` -- an agent that located a message via `mail.search` had no way to reply to it without an Outlook window focused and that exact item selected, and acted on whatever a human had selected instead of what the model intended (confused deputy, related to #9). `forward` also had no way to specify a recipient, so a forward draft had nobody to send to. All three now accept `entryId`/`storeId`/`useActiveMail` (routed through the same `OutlookInteropRunner.ResolveMailItem` used by every other mail action) and work correctly with no Outlook window focused when `entryId` is supplied. `forward` gained `recipientTo`/`cc`/`bcc`. All three gained an optional `body` parameter that is prepended above the quoted original message (Reply/ReplyAll/Forward pre-populate `Body` with the quoted original; prepending preserves that context rather than destroying it). Added integration smoke tests covering headless reply and forward-with-recipient.

- **`calendar` and `attachment` tools under-declared `destructiveHint`, `mail` over-declared it** (#18): action-dispatch tools (one MCP tool, many actions via an `action` enum) cannot be described by a single tool-level boolean. `calendar` declared `Destructive = false` while exposing `delete-appointment`/`update-appointment`; `attachment` declared `Destructive = false` while exposing `add`/`remove`; `mail` declared `Destructive = true` while exposing read-only `list`/`read`/`search`. Added a per-action `[ServiceAction(Destructive = ...)]` override, and the generator now computes each tool's `[McpServerTool(Destructive = ...)]` hint as true if ANY exposed action is destructive. `calendar`'s description no longer claims to manage appointments "safely". Added a regression test asserting every affected action's resolved classification and the generated tool-level hints.
- **Shared Outlook `Application` COM object could be invalidated process-wide** (#19): `OutlookInteropRunner` called `Marshal.FinalReleaseComObject` on the Outlook `Application` obtained from `GetActiveObject`, which is the user's already-running, shared instance cached per-process by the RCW table. Final-releasing it zeroed the ref-count for every holder in the process, risking `InvalidComObjectException` on subsequent operations. Now released via `Marshal.ReleaseComObject` (a plain ref-count decrement) instead, added a regression test that issues two sequential `Execute()` calls and asserts the second still succeeds.
- **`check-com-leaks.ps1` never scanned Outlook COM files** (#21): the script only flagged leaks in files using PowerPoint's `dynamic` COM pattern, so every Outlook file (which uses strongly-typed `Outlook.*` locals released via `OutlookInteropRunner.ReleaseComObject`) was silently skipped. The script now also detects the Outlook pattern and its release calls.
- **MCP `ServerInstructions` taught LLMs the legacy PowerPoint file/session workflow** (#24): the server-wide instructions sent to every connecting client told the model to unlock legacy session-gated tools via `file(action:'open')` → `session_id` → `file(action:'close')`, described the product as a "migration surface", and omitted `calendar` entirely. Rewritten to describe only the Outlook surface (`application`, `folder`, `mail`, `attachment`, `calendar`), the `entryId`/`storeId` identity model, when to use active-item targeting vs. explicit `entryId`, and destructive-action safety expectations.
- **Release workflow published under foreign package identities** (#5): removed the temporary `if: false` guard on the `publish` job now that every registry target is Outlook-owned.
- **MCP registry `server.json` version was never updated**: the version-rewrite regex matched a non-existent `Trsdn.PptMcp.McpServer` identifier and silently did nothing, so a stale version would have been published.
- **Scriban 6.6.0 broke every restore**: the templating engine behind SKILL.md generation carried a critical advisory (GHSA-5wr9-m6jw-xx44, patched in 7.0.0) and `NuGetAudit` treats it as an error, so `dotnet restore` failed outright. Bumped to 7.2.6; generated SKILL.md output is byte-identical.
- **Dependency review rejected the GitHub Copilot CLI license**: added an explicit `allow-dependencies-licenses` entry, since GitHub's proprietary terms cannot be expressed as an SPDX identifier.
- **Vulnerable transitive dependency in the agent lockfile**: `@github/copilot` resolved to 1.0.4, which is affected by GHSA-9ccr-r5hg-74gf (arbitrary command execution via `core.fsmonitor`). Refreshed to 1.0.80 within the existing `@github/copilot-sdk@0.1.32` range.
- **CI never ran on any branch**: all workflows filtered on `main`, but this repository's default branch is `master`, so build, CodeQL, dependency-review, and integration-test workflows were silently inert. `build-mcp-server.yml` was also missing a `pull_request` trigger entirely.
- **NuGet propagation check never succeeded**: the readme poll used a mixed-case package ID, which the lowercase-only flat-container API always answers with 404, wasting the full 30-minute retry window on every release.

### Added

- Official source-side Copilot SDK agent client under `src\OutlookMcp.Agent`, including local planner tests and documentation for the agent architecture
- Dedicated documentation for the evaluation framework and the archetype/reference pipeline
- **33 PowerPoint MCP tools with 204 operations** for comprehensive PowerPoint automation via COM interop
- **Slide management** (7 ops) — list, read, create, duplicate, move, delete, apply-layout
- **Shape operations** (17 ops) — add, move, resize, fill, line, shadow, rotation, z-order, grouping, copy between slides, connectors, merge shapes (union/combine/fragment/intersect/subtract)
- **Text editing** (6 ops) — get/set text, find, replace, format (font, size, bold, italic, color, alignment)
- **Charts** (5 ops) — create, get info, set title, set type, delete
- **Slide Tables** (9 ops) — create, read, write cells, add/delete rows and columns, merge cells
- **Animations** (4 ops) — list, add, remove, clear effects
- **Transitions** (3 ops) — get, set, remove slide transitions
- **Design/Themes** (4 ops) — list designs, apply themes, get theme colors, list color schemes
- **Images** (1 op) — insert with position and size control
- **Speaker Notes** (3 ops) — get, set, clear
- **Sections** (4 ops) — list, add, rename, delete presentation sections
- **Hyperlinks** (4 ops) — add, read, remove, list
- **Slideshow** (4 ops) — start, stop, navigate, get status
- **Slide Masters** (1 op) — list masters and layouts
- **Export** (4 ops) — PDF, slide images (PNG), video (MP4), print
- **VBA Macros** (5 ops) — list, view, import, delete, run
- **Media** (3 ops) — insert audio/video, get media info
- **Window Management** (4 ops) — get info, minimize, restore, maximize
- **File Validation** (1 op) — test file accessibility
- **Document Properties** (2 ops) — get/set title, author, subject, etc.
- **Comments** (4 ops) — list, add, delete, clear slide comments
- **Placeholders** (2 ops) — list placeholders, set placeholder text
- **Slide Background** (3 ops) — get info, set solid color, reset to master
- **Headers & Footers** (2 ops) — get/set footer text, slide numbers, date
- **SmartArt** (2 ops) — get diagram info, add nodes
- **Shape Alignment** (2 ops) — align and distribute shapes on slides
- **Custom Shows** (3 ops) — list, create, delete custom slide shows
- **Page Setup** (2 ops) — get/set slide size and orientation
- **Slide Import** (1 op) — import slides from another .pptx file
- **Tags** (3 ops) — custom metadata on slides and shapes
- **MCP Server** — Model Context Protocol server for AI assistants (GitHub Copilot, Claude, ChatGPT)
- **CLI** (`outlookcli`) — Command-line interface for scripting and coding agents
- **COM interop** — Uses PowerPoint's native COM API for 100% safe automation
- **Session management** — Shared sessions between MCP Server and CLI
- **Parameter validation** — All required string parameters validated before COM execution
- **COM resource safety** — All COM objects released in finally blocks to prevent leaks
