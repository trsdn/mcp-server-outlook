using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Signature;

/// <summary>
/// Pure filesystem logic for discovering Outlook signatures. Outlook stores each signature as up to
/// three sibling files sharing one base name — <c>Name.htm</c>, <c>Name.txt</c> and <c>Name.rtf</c> —
/// alongside a <c>Name_files</c> asset folder for the HTML version. Grouping those files back into
/// one signature per base name is the only non-trivial part, and it is entirely pure: given a folder
/// path it returns the grouped list without touching Outlook, COM, or any process state. It is
/// factored out here so it can be unit-tested against a temporary folder (see ADR-001).
/// </summary>
public static class SignatureFileScanner
{
    private static readonly string[] KnownExtensions = [".htm", ".html", ".txt", ".rtf"];

    /// <summary>
    /// Scans <paramref name="folderPath"/> for signature files and groups them by base name. Returns
    /// an empty list if the folder does not exist. Ordering is case-insensitive by name so the result
    /// is stable regardless of how the filesystem enumerates entries.
    /// </summary>
    public static List<OutlookSignatureInfo> Scan(string folderPath)
    {
        var byName = new Dictionary<string, OutlookSignatureInfo>(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            return [];
        }

        foreach (var file in Directory.EnumerateFiles(folderPath))
        {
            var extension = Path.GetExtension(file);
            if (!KnownExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var baseName = Path.GetFileNameWithoutExtension(file);
            if (string.IsNullOrEmpty(baseName))
            {
                continue;
            }

            if (!byName.TryGetValue(baseName, out var info))
            {
                info = new OutlookSignatureInfo { Name = baseName };
                byName[baseName] = info;
            }

            switch (extension.ToLowerInvariant())
            {
                case ".htm":
                case ".html":
                    info.HasHtml = true;
                    break;
                case ".txt":
                    info.HasText = true;
                    break;
                case ".rtf":
                    info.HasRtf = true;
                    break;
            }
        }

        return byName.Values
            .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// Resolves the file that holds a named signature in a given format, or null if it is not present.
    /// <paramref name="format"/> is one of <c>html</c>, <c>text</c>, <c>rtf</c>. Pure: no I/O beyond
    /// probing for the specific file's existence.
    /// </summary>
    public static string? ResolveSignatureFile(string folderPath, string name, string format)
    {
        if (string.IsNullOrWhiteSpace(folderPath) || string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        var extensions = format switch
        {
            "html" => new[] { ".htm", ".html" },
            "text" => [".txt"],
            "rtf" => [".rtf"],
            _ => System.Array.Empty<string>()
        };

        foreach (var extension in extensions)
        {
            var candidate = Path.Combine(folderPath, name + extension);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }
}
