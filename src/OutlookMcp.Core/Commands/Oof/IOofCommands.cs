using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Oof;

[ServiceCategory("oof")]
[McpTool("oof", Title = "Outlook Out-of-Office (Automatic Replies) Status", Destructive = false, Category = "settings",
    Description = "Read whether the mailbox's automatic replies (out-of-office) are currently ON or OFF. "
    + "READ-ONLY and PARTIAL: it reports only the on/off state, taken from the PR_OOF_STATE store "
    + "property, which is the single out-of-office facet classic Outlook exposes through COM. It CANNOT "
    + "read or set the reply message text, the separate internal and external replies, or a scheduled "
    + "start/end window - those require EWS or Microsoft Graph, not Outlook COM - and it cannot turn "
    + "OOF on or off. On a non-Exchange (POP/IMAP) store the feature does not apply. In Cached Exchange "
    + "mode the flag can lag the server until the next Send/Receive.")]
public interface IOofCommands
{
    /// <summary>
    /// Reports whether automatic replies (out-of-office) are currently enabled on the default Exchange
    /// mailbox, read from the <c>PR_OOF_STATE</c> store property via <c>Store.PropertyAccessor</c>.
    /// Read-only. Returns only the on/off boolean: the reply text, the separate internal and external
    /// variants, and any scheduled start/end window are NOT exposed through Outlook COM and require EWS
    /// or Microsoft Graph. For a non-Exchange store the feature does not apply and <c>isSupported</c> is
    /// false. In Cached Exchange mode the value can lag the server until the next synchronisation.
    /// </summary>
    [ServiceAction("get-status")]
    OutlookOofStatusResult GetStatus();
}
