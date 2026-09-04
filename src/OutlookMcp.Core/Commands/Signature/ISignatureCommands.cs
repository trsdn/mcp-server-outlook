using OutlookMcp.Core.Attributes;
using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Signature;

[ServiceCategory("signature")]
[McpTool("signature", Title = "Outlook Signature Operations", Destructive = false, Category = "settings",
    Description = "List and read the user's Outlook e-mail signatures. "
    + "NOTE: Outlook does not expose signatures through its COM object model; they are files under "
    + "%APPDATA%\\Microsoft\\Signatures. These actions therefore read that folder directly and are "
    + "strictly read-only — they never create, change, delete, or apply a signature, and they cannot "
    + "set the signature Outlook uses for new mail (that is a per-account setting the object model "
    + "does not expose). Use list to discover signature names and which formats (html/text/rtf) each "
    + "has, and read to fetch the content of one — for example to append it to a draft you are "
    + "composing.")]
public interface ISignatureCommands
{
    /// <summary>
    /// Lists the signatures available to the user by scanning <c>%APPDATA%\Microsoft\Signatures</c>.
    /// Read-only, filesystem only — no Outlook or COM involvement. If the folder does not exist (the
    /// user has never created a signature), this succeeds with an empty list and
    /// <c>folderExists=false</c> rather than failing.
    /// </summary>
    [ServiceAction("list")]
    OutlookSignatureListResult ListSignatures();

    /// <summary>
    /// Reads the content of one signature in a chosen format. Read-only, filesystem only.
    /// </summary>
    /// <param name="signatureName">The signature name exactly as returned by list (the file name without its extension).</param>
    /// <param name="format">Which rendering to return: text (default, plain text), html, or rtf. If the requested format is not present for that signature, this is an error that reports which formats are available.</param>
    [ServiceAction("read")]
    OutlookSignatureReadResult ReadSignature(string signatureName, string format = "text");
}
