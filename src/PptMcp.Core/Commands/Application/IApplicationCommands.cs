using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Application;

[ServiceCategory("application")]
[NoSession]
[McpTool("application", Title = "Outlook Application Operations", Destructive = false, Category = "application",
    Description = "Inspect the current Outlook desktop application state without opening a persistent session. "
    + "Use get-status to verify Outlook availability, inspect active explorer/inspector counts, and read the current folder context.")]
public interface IApplicationCommands
{
    [ServiceAction("get-status")]
    OutlookApplicationStatusResult GetStatus(bool includeActiveContext = true);
}
