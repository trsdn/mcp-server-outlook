## Summary
Brief description of what this PR does.

## Type of Change
- [ ] Bug fix (non-breaking change which fixes an issue)
- [ ] New feature (non-breaking change which adds functionality)
- [ ] Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] Documentation update
- [ ] Maintenance (dependency updates, code cleanup, etc.)

## Related Issues
Closes #[issue number]
Relates to #[issue number]

## Changes Made
- Change 1
- Change 2

## Testing Performed

> **Read this before ticking anything.** No CI job verifies Outlook behaviour, and none ever will:
> there is no self-hosted Windows runner with classic Outlook, and the repository owner has decided
> there will not be one (#31, closed as not planned). Every Outlook claim below therefore rests on a
> **local** run against a real profile. Say so explicitly rather than implying CI covered it, and if
> you did not run it locally, say that instead of ticking the box.

- [ ] Build produces zero warnings and zero errors
- [ ] Ran the targeted test filter for the area I changed (not the full suite)
- [ ] Tested error conditions (Outlook not running, new-Outlook-only, stale entry ID, invalid arguments)
- [ ] Verified COM objects are released: `scripts\check-com-leaks.ps1` reports 0 leaks
- [ ] Manually exercised the change against a real Outlook profile, and said which actions below

**Manual verification performed (be specific):**
```powershell
# Actual commands run, and what you observed
outlookcli application get-status
```

## New Action Checklist

**Does this PR add or modify a Core Commands method?** [ ] Yes [ ] No

If YES, every sync point must be updated. The MCP server and the CLI are both first-class surfaces
and are generated from the same interfaces, so a half-finished action ships as broken parity.

- [ ] Added the method to the Core Commands interface (e.g. `IMailCommands`)
- [ ] Implemented it in the Core Commands class (e.g. `MailCommands`)
- [ ] Annotated the interface method with `[ServiceAction]`, and `Destructive = true` if it mutates or deletes
- [ ] Ran `CoreCommandsCoverageTests` to confirm the generated enum matches the interface
- [ ] Verified the action appears in **both** the generated MCP tool and the generated CLI command
- [ ] Added integration tests
- [ ] Updated `FEATURES.md`, including the operation count
- [ ] Updated `skills/shared/*.md` if the guidance changed (these become MCP prompts)
- [ ] Updated the README operation counts if they changed

## Safety Review

- [ ] This PR does not add a way to send or delete mail without explicit confirmation
- [ ] Any new irreversible action is idempotent per operation ID, or explains why it need not be
- [ ] No mailbox content, entry IDs, addresses, or file paths from a real mailbox appear in the diff,
      tests, commit messages, or this description

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review completed
- [ ] `Success` is never `true` alongside an `ErrorMessage`
- [ ] Core commands do not catch exceptions inside the action lambda; failures flow through the runner's `onException`
- [ ] Every COM object is released in a `finally` block via `OutlookInteropRunner.ReleaseComObject`, and the shared `Outlook.Application` is never final-released
- [ ] No TODO / FIXME / HACK markers, and no commented-out code
- [ ] Escapes user input with `.EscapeMarkup()` in CLI output
- [ ] Returns consistent exit codes (`0` on success, non-zero when the operation reports `success: false`)
- [ ] CHANGELOG.md updated under `## [Unreleased]` for any user-visible change
- [ ] No confidential or personal information in the commits, description, or tests

## Additional Notes
Anything reviewers should know, including trade-offs and areas you would like careful review on.
