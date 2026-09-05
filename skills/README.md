# Outlook agent skills

Two skill packages target the two entry points. Both describe the same 10 tools and 69 operations.

| Skill | Folder | Target | Best for |
|---|---|---|---|
| `outlook-cli` | `skills/outlook-cli/` | The `outlookcli` CLI | Coding agents; token-efficient and `--help` discoverable |
| `outlook-mcp` | `skills/outlook-mcp/` | The MCP server | Conversational AI; rich tool schemas |

## Single source of truth

`skills/shared/*.md` is authoritative. On a Release build it is:

1. copied into `skills/outlook-cli/references/` and `skills/outlook-mcp/references/`, for skill-based clients such as VS Code and Cursor
2. code-generated into `[McpServerPrompt]` methods, so MCP-only clients such as Claude Desktop get the same guidance

Never create a separate prompt file for content that belongs in `skills/shared/`. Every file in that
directory becomes an MCP prompt, so it must be Outlook guidance.

`SKILL.md` in each package is **generated** from `skills/templates/*.sbn` plus the generated skill
manifest. Do not hand-edit it; a Release build overwrites it. Edit the template instead.

## Install

```powershell
npx skills add trsdn/mcp-server-outlook --skill outlook-cli   # coding agents
npx skills add trsdn/mcp-server-outlook --skill outlook-mcp   # conversational AI
```

## Naming note

Project and type names are still inherited: the solution uses `OutlookMcp.*`, and some
infrastructure is now Outlook-named (see #12). The skill folders themselves are Outlook-named
and describe only Outlook behaviour.
