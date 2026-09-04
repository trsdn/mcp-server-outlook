# OutlookMcp.CLI

The `outlookcli` command-line surface for Outlook automation.

The CLI and the MCP server are **both first-class entry points**. They are source-generated from the
same `[ServiceCategory]` interfaces in `OutlookMcp.Core`, so they expose identical actions,
parameters, defaults, and validation.

## Surface

$18 tools, 62 actions:

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
