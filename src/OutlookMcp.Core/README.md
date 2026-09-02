# OutlookMcp.Core

Outlook business logic for OutlookMcp. Provides commands for mail, calendar, folders, attachments,
and application status.

This is an internal library used by
[OutlookMcp.McpServer](https://www.nuget.org/packages/OutlookMcp.McpServer) and
[OutlookMcp.CLI](https://www.nuget.org/packages/OutlookMcp.CLI). It is not intended for direct
consumption.

Command interfaces here are marked `[ServiceCategory]`, and both the MCP tool surface and the CLI
command surface are source-generated from them. That is what keeps the two entry points in parity:
adding an action to an interface adds it to both.

## Requirements

- Windows
- .NET 9.0+
- The **classic Outlook for Windows desktop app**, installed and running. The new Outlook for
  Windows exposes no COM object model and cannot be automated.

## Documentation

See the [main repository](https://github.com/trsdn/mcp-server-outlook) for full documentation.
