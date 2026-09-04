# OutlookMcp.CLI

The `outlookcli` command-line surface for Outlook automation.

The CLI and the MCP server are **both first-class entry points**. They are source-generated from the
same `[ServiceCategory]` interfaces in `OutlookMcp.Core`, so they expose identical actions,
parameters, defaults, and validation.

## Surface

5 tools, 30 actions:

| Tool | Actions |
|---|---|
| `mail` | `read-active`, `read`, `list`, `search`, `create-draft`, `reply`, `reply-all`, `forward`, `send`, `move`, `delete`, `set-read-state`, `set-categories`, `set-subject`, `set-body`, `set-recipients` |
| `calendar` | `list`, `read`, `create-appointment`, `update-appointment`, `delete-appointment` |
| `folder` | `list-default`, `list-children`, `resolve-path`, `list-items` |
| `attachment` | `list`, `save`, `add`, `remove` |
| `application` | `get-status` |

See [FEATURES.md](../../FEATURES.md) for descriptions.

## Usage

```powershell
outlookcli application get-status
outlookcli folder list-default
outlookcli mail list --folder Inbox
```

Everything is discoverable through `--help`:

```powershell
outlookcli --help
outlookcli mail --help
outlookcli mail search --help
```

This makes the CLI the better surface for coding agents: it is token-efficient and self-describing,
where the MCP server pays for richer schemas up front.

## Requirements

- Windows
- The **classic Outlook for Windows desktop app**, installed and running. Run
  `outlookcli application get-status` first; if it reports `NewOutlookOnly`, the new Outlook for
  Windows is the only client present and cannot be automated.

## Optional: restrict who `mail send` may write to

Set `OUTLOOKMCP_ALLOWED_RECIPIENTS` to a semicolon- or comma-separated allow-list of permitted
recipient domains and addresses:

```powershell
$env:OUTLOOKMCP_ALLOWED_RECIPIENTS = "@contoso.example; alice@fabrikam.example"
outlookcli mail send --entry-id $id --confirm true
```

`mail send` then refuses any recipient outside the list before Outlook is asked to send anything. A
bare domain and an `@`-prefixed domain mean the same; an entry with a local part is an exact
address. Matching is case-insensitive, a domain entry does not admit its subdomains, and an address
that cannot be read as SMTP is refused rather than assumed safe.

**The variable is unset by default and the feature is off.** Set it before the daemon starts - it is
read per send, but the daemon inherits the environment of whichever `outlookcli` invocation
launched it. The same variable configures the MCP server, via the `env` block of its client
definition.

## Exit codes

- `0` - the command ran and the operation succeeded
- `1` - the arguments could not be parsed, or the operation reported `success: false`

The exit code reflects the operation, not just the transport, so it is safe to branch on
(this was not true before #63):

```powershell
outlookcli folder list-default
if ($LASTEXITCODE -ne 0) { throw "folder list-default failed" }
```

Commands whose payload has no `success` property (bare arrays and ordinary read results) exit `0`.
The JSON is still authoritative if you need the error text:

```powershell
$r = outlookcli folder list-default | ConvertFrom-Json
if (-not $r.success) { throw $r.errorMessage }
```

When `--output <path>` is given and the operation fails, no file is written and the error JSON goes
to stdout.

## Naming note

Project and assembly names are still inherited (`OutlookMcp.*`), and some shared infrastructure is
now Outlook-named throughout. Any residual naming debt is tracked as #12. It is transitional, not the intended
long-term branding.
