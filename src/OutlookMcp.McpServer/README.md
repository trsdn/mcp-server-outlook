# OutlookMcp.McpServer

The Model Context Protocol server for Outlook automation.

The MCP server and the `outlookcli` CLI are **both first-class entry points**. They are
source-generated from the same `[ServiceCategory]` interfaces in `OutlookMcp.Core`, so they expose
identical actions, parameters, defaults, and validation.

## Tool families

9 tools, 64 operations:

| Tool | Operations |
|---|---|
| `mail` | 23 |
| `folder` | 10 |
| `calendar` | 7 |
| `contact` | 5 |
| `task` | 5 |
| `attachment` | 4 |
| `application` | 3 |
| `addressbook` | 3 |
| `property` | 4 |

See [FEATURES.md](../../FEATURES.md) for the full action list.

There are no `slide`, `shape`, `text`, `chart`, `animation`, `transition`, `slideshow`, `file`, or
`window` tools. Those inherited families were deleted in #26.

## Prompts

The server ships MCP prompts generated from `skills/shared/*.md`, so MCP-only clients such as Claude
Desktop receive the same behavioural guidance that skill-based clients read from the skill packages.
`skills/shared/` is the single source of truth; every `.md` file there becomes a prompt, so it must
contain Outlook guidance only.

## Requirements

- Windows
- .NET 9.0+
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows exposes no COM object model and cannot be automated; `application.get-status` reports
  `NewOutlookOnly` in that case.

Have the assistant call `application.get-status` first to confirm the environment.

## Client configuration

See [`examples/mcp-configs/`](../../examples/mcp-configs/) for Claude Desktop, VS Code, Cursor,
Cline, and Windsurf snippets, and [`mcpb/`](../../mcpb/) for the Claude Desktop bundle.

## Naming note

Project and assembly names are `OutlookMcp.*`, the hand-written base class is `OutlookToolsBase`
and the generated tool types are `Outlook*Tool`. Any residual naming debt is tracked as
#12. It is transitional, not the intended long-term branding.
