# outlook-cli-skill

An [Agent Skill](https://agentskills.io) for automating Microsoft PowerPoint via the [outlookcli](https://github.com/trsdn/mcp-server-outlook) command-line tool.

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent how to automate PowerPoint from scripts and CI/CD pipelines:

- **Presentation management** — open, create, save, close
- **Range operations** — read/write values, formatting, formulas
- **Tables & PivotTables** — create, modify, refresh
- **Charts** — create and configure chart types
- **Power Query (M code)** — create and edit queries
- **Data Model (DAX)** — add measures and calculated columns
- **VBA macros, conditional formatting**, and more

## Requirements

- Windows with Microsoft PowerPoint 2016+ installed
- Install the CLI: `dotnet tool install --global OutlookMcp.CLI`

## Install

```bash
npx skillpm install outlook-cli-skill
```

Or with npm directly:

```bash
npm install outlook-cli-skill
```

## License

MIT
