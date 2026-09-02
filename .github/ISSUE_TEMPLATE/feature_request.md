---
name: Feature Request
about: Suggest an idea for OutlookMcp
title: '[FEATURE] '
labels: 'enhancement'
assignees: ''

---

## Is your feature request related to a problem?
A clear and concise description of the problem. Ex. I'm always frustrated when [...]

## Component
Which component should this feature be added to?
- [ ] **MCP Server** (AI assistant integration)
- [ ] **CLI** (`outlookcli`)
- [ ] **Both** (features must reach both surfaces; they are generated from the same interfaces)
- [ ] **Core Library** (shared functionality)
- [ ] **VS Code extension**
- [ ] **Not sure**

## Describe the solution you'd like
A clear and concise description of what you want to happen.

## Proposed Syntax

**For CLI:**
```powershell
outlookcli <tool> <new-action> [options]
```

**For MCP Server:**
- Tool: [one of the existing tools: mail, calendar, folder, attachment, application - or propose a new one]
- Action: [e.g. new-action]
- Parameters: [describe expected parameters]

Note that the MCP and CLI surfaces are source-generated from the same `[ServiceCategory]` interfaces
in Core, so a new action lands on both automatically. Please do not propose a feature for only one.

## Describe alternatives you've considered
Any alternative solutions or features you have considered.

## Outlook Domain
Which part of Outlook does this touch?
- [ ] Mail (read, list, search, draft, reply, send, move, flag, categorise)
- [ ] Calendar (appointments, meetings, attendees)
- [ ] Folders (navigation, resolution, item listing)
- [ ] Attachments (list, save, add, remove)
- [ ] Contacts (**not currently exposed at all**)
- [ ] Tasks / follow-up (**not currently exposed at all**)
- [ ] Rules, categories, or account configuration (**not currently exposed at all**)
- [ ] Other: [please specify]

## Target Users
Who would benefit?
- [ ] **AI assistants** (GitHub Copilot, Claude Desktop, via the MCP Server)
- [ ] **Direct CLI users** (scripted automation)
- [ ] **Coding agents** (token-efficient CLI surface)
- [ ] Other: [please specify]

## Safety Considerations

Outlook automation acts on a real mailbox and is often irreversible.

- [ ] This feature can **send** mail
- [ ] This feature can **delete** items
- [ ] This feature modifies items the user did not explicitly name
- [ ] This feature is read-only

If any of the first three are checked, say what confirmation or idempotency the action should have.
For reference, `mail.send` requires explicit confirmation and is idempotent per operation ID.

## Additional context
Any other context or examples.

## Implementation Notes
If you have ideas about how this could be implemented, share them here. Note that Outlook COM runs
through `OutlookDispatcher` on a dedicated STA thread; see
[ADR-002](../../docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md).
