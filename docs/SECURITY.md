# Security Policy

## Supported Versions

We actively support the current major version of OutlookMcp with security updates.

| Version | Supported | Status |
| ------- | --------- | ------ |
| 1.x.x   | Yes       | Active |

## Security Features

OutlookMcp includes security measures appropriate for local Outlook COM automation.

### Input Validation

- File paths used for attachments and output files are normalized and validated by the command layer.
- Destructive operations require explicit action names and parameters.
- `mail send` requires `confirm=true` before a draft is sent.
- Optional `operationId` support on `mail send` reduces duplicate-send risk after retries.

### Code Analysis

- Security and reliability analyzers run as part of the build.
- Warnings are treated as errors.
- CodeQL and dependency scanning are configured in repository workflows.

### Outlook COM Security

- OutlookMcp uses the classic Outlook for Windows COM object model through `Outlook.Application`.
- New Outlook for Windows is not supported because it exposes no COM object model.
- Outlook must already be running. The command layer deliberately avoids creating a new untrusted Outlook instance for normal operations.
- Operations run as the current Windows user and do not elevate privileges.
- Outlook Object Model Guard prompts can block sensitive operations. A person must respond in Outlook; OutlookMcp cannot bypass or approve those prompts.
- Entry IDs identify items, but can change when an item moves between stores.

### OutlookMcp Service Security

The MCP server runs Outlook operations in-process. The CLI uses a local daemon and named pipe.

**MCP Server:** Runs fully in-process with the MCP host. There is no additional network listener created by OutlookMcp.

**CLI:** Uses a Windows named pipe for communication between `outlookcli` commands and the daemon process.

| Protection | Status | Description |
| ---------- | ------ | ----------- |
| User isolation | Enforced | The pipe name includes the current Windows user SID. |
| Windows ACLs | Enforced | The named pipe restricts access to the current user. |
| Local only | Enforced | Named pipes are local IPC only. |
| Same-user process restriction | Not enforced | Any same-user process can connect to the daemon. |

What this means:

1. Same-user applications can request Outlook operations through the CLI daemon.
2. Other Windows users cannot access the daemon through the default pipe.
3. Remote network clients cannot connect to the daemon.
4. The daemon does not provide capabilities beyond what the current user can already do in Outlook.

### Dependency Management

- Dependabot monitors package updates.
- Dependency review checks pull requests.
- Central package management keeps versions consistent.

## Reporting a Vulnerability

We take security vulnerabilities seriously. If you discover a security issue, please follow these steps.

### 1. Do Not Create a Public Issue

Please do not create a public GitHub issue for security vulnerabilities.

### 2. Report Privately

Preferred method: GitHub Security Advisories

1. Go to <https://github.com/trsdn/mcp-server-outlook/security/advisories>.
2. Click "Report a vulnerability".
3. Fill out the advisory form.

Alternative: contact the maintainer through GitHub at [@trsdn](https://github.com/trsdn).

Subject: `[SECURITY] OutlookMcp Vulnerability Report`

### 3. Information to Include

- Description: Clear description of the vulnerability.
- Impact: What could an attacker do?
- Affected Versions: Which versions are affected?
- Proof of Concept: Steps to reproduce, if possible.
- Suggested Fix: Optional mitigation or patch idea.

Example:

```text
Vulnerability: Unsafe attachment save path handling
Impact: Attacker could write files outside the intended directory
Affected Versions: 1.0.0 - 1.0.2
PoC: outlookcli attachment save --destination-directory "..\..\target"
Suggested Fix: Validate normalized output paths before writing
```

### 4. What to Expect

- Acknowledgment within 48 hours.
- Initial assessment within 5 business days.
- Regular status updates.
- Fix timeline based on severity.

### 5. Coordinated Disclosure

1. Private fix development.
2. GitHub Security Advisory when appropriate.
3. CVE request when applicable.
4. Public release with security notes.
5. Researcher credit if desired.

## Security Best Practices for Users

### MCP Server Security

- Use only trusted MCP hosts and AI assistants.
- Review destructive tool calls before allowing them.
- Remember that output can contain mailbox data.
- Monitor MCP client logs according to your local policy.

### CLI Security

- Review scripts before running them.
- Protect files written by attachment save operations.
- Avoid running the CLI as Administrator unless Outlook is also running at the same elevation level.
- Stop the CLI daemon when you do not need it:

```powershell
outlookcli service stop
```

### Outlook Data Security

- Be cautious with attachment add and save operations.
- Re-read or re-list items after moving them between stores because entry IDs can change.
- Treat sent mail, deleted mail, moved mail, and calendar updates as real mailbox changes.

### Development Security

- All changes require pull request review before merge.
- Main branch protection and CI checks are expected.
- Build with zero warnings.
- Do not include private mailbox data, customer names, or secrets in tests, issues, logs, or PRs.

## Known Security Considerations

### Classic Outlook Dependency

- Local only.
- Windows only.
- Requires classic Outlook for Windows desktop app.
- Does not support new Outlook for Windows.
- Uses the user's already-running Outlook session.

### Shared Desktop Application

- Outlook is a single shared desktop app per user session.
- Automation is not isolated from the Outlook instance the user is also operating.
- There is no supported show or hide control for Outlook windows.

### AI Integration

- Only use trusted AI platforms.
- Review operations that send, move, delete, or mutate mailbox data.
- Avoid exposing sensitive mailbox data to AI assistants unless your policy allows it.

### Testing Gap

No Outlook behavior is currently verified by hosted CI. Outlook integration tests require a self-hosted Windows runner with classic Outlook installed and signed in.

## Security Updates

Security updates are published through:

- GitHub Security Advisories: <https://github.com/trsdn/mcp-server-outlook/security/advisories>
- Release Notes: <https://github.com/trsdn/mcp-server-outlook/releases>
- NuGet Advisories: package vulnerabilities shown in NuGet

Subscribe to repository notifications to receive security alerts.

## Vulnerability Disclosure Policy

### Our Commitment

- We will acknowledge receipt of vulnerability reports within 48 hours.
- We will keep reporters informed of progress.
- We will credit researchers in security advisories if desired.
- We will not take legal action against researchers following responsible disclosure.

### Researcher Guidelines

- Give us time to fix before public disclosure.
- Do not access, modify, or delete other users' data.
- Act in good faith to help improve security.
- Follow all applicable laws.

## Security Contacts

- GitHub Security: <https://github.com/trsdn/mcp-server-outlook/security>
- Maintainer: @trsdn

## Additional Resources

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Microsoft Security Response Center](https://msrc.microsoft.com/)
- [CVE Database](https://cve.mitre.org/)
- [National Vulnerability Database](https://nvd.nist.gov/)

## Version History

| Version | Date | Security Changes |
| ------- | ---- | ---------------- |
| 1.x     | 2026 | Outlook COM command surface with STA dispatcher and CLI named pipe isolation |

---

**Last Updated:** 2026-09-02

Thank you for helping keep OutlookMcp and its users safe.
