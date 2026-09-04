using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Application;

[ServiceCategory("application")]
[McpTool("application", Title = "Outlook Application Operations", Destructive = false, Category = "application",
    Description = "Inspect the current Outlook desktop application state without opening a persistent session. "
    + "Use get-status to verify Outlook availability, inspect active explorer/inspector counts, and read the current folder context. "
    + "Use get-active-explorer to find out which folder the user is looking at and what is selected there, and "
    + "get-active-inspector to find out which item the user currently has open. "
    + "Both answer \"nothing is open\" as a success rather than an error. "
    + "An item the user is still composing has not been saved and therefore has no entryId, so it cannot be "
    + "addressed by any other action until it is saved; isSaved says which case you are in.")]
public interface IApplicationCommands
{
    [ServiceAction("get-status")]
    OutlookApplicationStatusResult GetStatus(bool includeActiveContext = true);

    [ServiceAction("get-active-explorer")]
    OutlookExplorerContextResult GetActiveExplorer();

    [ServiceAction("get-active-inspector")]
    OutlookInspectorContextResult GetActiveInspector();
}
