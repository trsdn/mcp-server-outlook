using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.OutlookInterop;

/// <summary>
/// Everything about an export destination that can be decided before a COM object exists, shared by
/// mail and calendar export (#14).
///
/// <para>
/// <c>SaveAs</c> looks trivial and is not. Four behaviours were measured against real Outlook before
/// any of this was written, and each one hands the caller a confidently wrong result if passed
/// straight through:
/// </para>
///
/// <list type="bullet">
///   <item><description>
///     <b>olMSG is silently lossy.</b> A subject of <c>probe Grüsse äöü тест €</c> saved with
///     <c>olMSG</c> (3) and reopened reads <c>probe Grüsse äöü ???? €</c>. The Cyrillic is gone and
///     nothing reports it. <c>olMSGUnicode</c> (9) round-trips exactly, so <c>msg</c> means Unicode
///     here and the ANSI constant is never used.
///   </description></item>
///   <item><description>
///     <b>The extension is ignored.</b> <c>SaveAs("x.txt", olMSGUnicode)</c> writes an OLE compound
///     file to a <c>.txt</c> path without complaint.
///   </description></item>
///   <item><description>
///     <b>A relative path succeeds.</b> It resolves against Outlook's working directory, not the
///     caller's, so the file exists somewhere nobody will look.
///   </description></item>
///   <item><description>
///     <b>A missing directory reports "The operation failed."</b>, which names nothing.
///   </description></item>
/// </list>
///
/// <para>
/// Planning happens before the item is resolved so that a rejected export cannot leave a partly
/// written file next to a failure result.
/// </para>
/// </summary>
internal static class ItemExportPlanner
{
    /// <summary>
    /// Formats <c>SaveAs</c> can produce, keyed by the extension that names them. The MSG entry
    /// deliberately maps to <c>olMSGUnicode</c> rather than <c>olMSG</c>; see the class remarks.
    /// </summary>
    private static readonly Dictionary<string, (string Format, Outlook.OlSaveAsType Type)> Formats =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["msg"] = ("msg", Outlook.OlSaveAsType.olMSGUnicode),
            ["txt"] = ("txt", Outlook.OlSaveAsType.olTXT),
            ["html"] = ("html", Outlook.OlSaveAsType.olHTML),
            ["htm"] = ("html", Outlook.OlSaveAsType.olHTML),
            ["mht"] = ("mht", Outlook.OlSaveAsType.olMHTML),
            ["mhtml"] = ("mht", Outlook.OlSaveAsType.olMHTML),
            ["rtf"] = ("rtf", Outlook.OlSaveAsType.olRTF),
            ["ics"] = ("ics", Outlook.OlSaveAsType.olICal),
        };

    /// <summary>The formats a caller may name, for use in error messages.</summary>
    internal static string SupportedFormats { get; } = string.Join(
        ", ",
        Formats.Values.Select(v => v.Format).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal));

    internal static bool TryPlan(
        string filePath,
        string? format,
        bool overwrite,
        out ExportPlan plan,
        out ItemExportResult? error)
    {
        plan = default;
        error = null;

        if (string.IsNullOrWhiteSpace(filePath))
        {
            error = Fail("filePath is required.");
            return false;
        }

        // Outlook accepts a relative path and resolves it against its own working directory, so the
        // export succeeds and lands somewhere the caller never looks. Refuse instead.
        if (!Path.IsPathRooted(filePath))
        {
            error = Fail(
                $"filePath must be an absolute path. '{filePath}' is relative, and Outlook would "
                + "resolve it against its own working directory rather than yours.");
            return false;
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(filePath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            error = Fail($"filePath '{filePath}' is not a usable path: {ex.Message}");
            return false;
        }

        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory))
        {
            error = Fail($"filePath '{filePath}' has no directory to write into.");
            return false;
        }

        // Outlook's own message for this is "The operation failed.", which names nothing.
        if (!Directory.Exists(directory))
        {
            error = Fail($"The directory '{directory}' does not exist. Create it first, or export elsewhere.");
            return false;
        }

        string extension = Path.GetExtension(fullPath).TrimStart('.');
        bool haveExtensionFormat = Formats.TryGetValue(extension, out var fromExtension);

        string resolvedFormat;
        Outlook.OlSaveAsType resolvedType;

        if (string.IsNullOrWhiteSpace(format))
        {
            if (!haveExtensionFormat)
            {
                error = Fail(
                    $"Cannot tell what format '{Path.GetFileName(fullPath)}' should be. Pass format "
                    + $"explicitly, or use one of these extensions: {SupportedFormats}.");
                return false;
            }

            (resolvedFormat, resolvedType) = fromExtension;
        }
        else
        {
            string requested = format.Trim().TrimStart('.');
            if (!Formats.TryGetValue(requested, out var explicitFormat))
            {
                error = Fail($"Unknown format '{format}'. Supported: {SupportedFormats}.");
                return false;
            }

            (resolvedFormat, resolvedType) = explicitFormat;

            // Outlook writes the requested format whatever the extension says, so allowing a
            // mismatch produces a file every other tool will misread - a binary .msg under a .txt
            // name. Refuse rather than mislabel.
            if (haveExtensionFormat && !string.Equals(fromExtension.Format, resolvedFormat, StringComparison.Ordinal))
            {
                error = Fail(
                    $"format '{resolvedFormat}' contradicts the '.{extension}' extension. Outlook "
                    + "writes the requested format regardless of the name, so this would produce a "
                    + $"{resolvedFormat} file called '{Path.GetFileName(fullPath)}'. Rename the file "
                    + "or drop the format argument.");
                return false;
            }
        }

        // Outlook overwrites silently, and an export that destroys the previous export without
        // saying so is not recoverable.
        bool exists = File.Exists(fullPath);
        if (exists && !overwrite)
        {
            error = Fail($"'{fullPath}' already exists. Pass overwrite to replace it.");
            return false;
        }

        plan = new ExportPlan(fullPath, resolvedFormat, resolvedType, exists);
        return true;

        static ItemExportResult Fail(string message) => new() { Success = false, ErrorMessage = message };
    }

    /// <summary>
    /// Reads back what actually landed on disk. Outlook reports nothing about what it wrote, and an
    /// export that reports success while writing no file is exactly the failure this project keeps
    /// finding elsewhere.
    /// </summary>
    internal static ItemExportResult? VerifyWritten(in ExportPlan plan)
    {
        var written = new FileInfo(plan.FilePath);
        return written.Exists
            ? null
            : new ItemExportResult
            {
                Success = false,
                ErrorMessage = $"Outlook reported no error but wrote no file at '{plan.FilePath}'."
            };
    }
}

internal readonly record struct ExportPlan(
    string FilePath,
    string Format,
    Outlook.OlSaveAsType Type,
    bool Overwriting);
