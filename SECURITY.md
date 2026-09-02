# Security Policy

## Supported Versions

We currently support the current major version of OutlookMcp with security updates.

| Version | Supported |
| ------- | --------- |
| 1.x.x   | Yes       |

## Reporting a Vulnerability

We take security seriously. If you discover a security vulnerability in OutlookMcp, please report it responsibly.

### How to Report

1. Do not create a public GitHub issue for security vulnerabilities.
2. Use GitHub Security Advisories for this repository when possible.
3. Include the following information:
   - Description of the vulnerability.
   - Steps to reproduce the issue.
   - Potential impact.
   - Suggested fix, if you have one.

### What to Expect

- We will acknowledge receipt of your vulnerability report within 48 hours.
- We will provide an estimated timeline for addressing the vulnerability within 1 week.
- We will notify you when the vulnerability has been fixed.
- We will credit you in the security advisory if you wish.

## Security Considerations

### Outlook COM Automation

OutlookMcp automates the classic Outlook for Windows desktop app through the `Outlook.Application` COM ProgID.

Important boundaries:

- New Outlook for Windows is not supported because it has no COM object model.
- OutlookMcp acts on the user's real mailbox and the same Outlook desktop app the user is using.
- Automation is not isolated from the interactive desktop session.
- There is no Outlook window show or hide control.
- The server does not elevate privileges. Operations run as the current Windows user.
- Outlook Object Model Guard prompts can appear for sensitive operations such as reading protected address fields or sending mail. OutlookMcp cannot approve those prompts automatically.

### Entry IDs and Mailbox State

Outlook items are addressed by entry ID, usually obtained from `mail list`, `mail search`, `mail read-active`, `calendar list`, or similar results.

- Entry IDs can change when items move between stores.
- Destructive operations such as send, move, delete, appointment update, appointment delete, attachment add, and attachment remove affect real Outlook data.
- `mail send` requires `confirm=true` and supports an optional `operationId` for duplicate-send protection on retries.

### Execution Model

Outlook COM work is serialized through `OutlookDispatcher` on a dedicated STA thread. See `docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md` for the documented model.

The retained `PptSession`, `PptBatch`, `PptContext`, `SessionManager`, and `OleMessageFilter` layer is dormant legacy infrastructure. It is not used by the current Outlook command surface.

### CLI Daemon

The CLI uses a local Windows named pipe for communication with its daemon process.

| Protection | Status | Description |
| ---------- | ------ | ----------- |
| User isolation | Enforced | The default pipe name includes the current user's SID. |
| Windows ACLs | Enforced | Pipe access is restricted to the current user. |
| Local only | Enforced | The named pipe is local IPC, not a network endpoint. |
| Same-user restriction | Not enforced | Any process running as the same Windows user can request CLI operations. |

This does not grant permissions beyond the current user, but it matters if untrusted software is already running as that user.

### Dependency Security

- .NET 10 is Microsoft-maintained and receives security updates.
- Dependencies are managed centrally.
- Dependabot and dependency review help identify vulnerable packages.
- The project does not require external cloud services for core Outlook automation.

### Best Practices for Users

1. Use OutlookMcp only with trusted MCP clients and trusted automation scripts.
2. Review AI-requested destructive operations before approving them.
3. Keep classic Outlook, Windows, and .NET updated.
4. Run OutlookMcp with the least necessary privileges.
5. Avoid running Outlook elevated unless the MCP server or CLI is also elevated for a specific reason.
6. Re-check entry IDs after moving items between stores.
7. Be cautious when saving or adding attachments from untrusted sources.
8. Treat command output as potentially sensitive because it can include mail, calendar, and folder metadata.

### Known Limitations

- Windows only.
- Requires classic Outlook for Windows desktop app.
- Does not support new Outlook for Windows.
- Acts on the user's real mailbox.
- No Outlook behavior is currently verified by hosted CI because integration testing requires a self-hosted Windows runner with classic Outlook.

## Version Updates

- Security patches will be released as soon as possible.
- Users are encouraged to keep OutlookMcp updated to the latest version.
- Breaking changes will be documented in release notes.

## Contact

For security-related questions or concerns, use the security reporting method above. For non-sensitive matters, use GitHub issues.
