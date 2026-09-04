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

        // Security: a signature name is a bare file base name and nothing more. Reject anything that
        // could escape the signatures folder before it ever reaches the filesystem. This matters
        // because this server reads e-mail, and an MCP client can be fed untrusted content from a
        // message body; without this guard a crafted signatureName is a path to an arbitrary
        // .htm/.html/.txt/.rtf file read.
        //   - A ROOTED name is the nastiest case: Path.Combine(base, rooted) discards base entirely
        //     and returns the rooted path, so "C:\some\secret" would read C:\some\secret.txt.
        //   - A name containing a directory separator (e.g. "..\..\etc") enables classic traversal.
        if (Path.IsPathRooted(name) ||
            name.Contains(Path.DirectorySeparatorChar) ||
            name.Contains(Path.AltDirectorySeparatorChar))
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

        // Canonicalise the folder once so the candidate can be verified to live directly under it.
        string folderFull = Path.GetFullPath(folderPath);
        string folderPrefix = folderFull.EndsWith(Path.DirectorySeparatorChar)
            ? folderFull
            : folderFull + Path.DirectorySeparatorChar;

        foreach (var extension in extensions)
        {
            string candidate = Path.GetFullPath(Path.Combine(folderFull, name + extension));

            // Belt-and-braces: even if some exotic name slipped past the checks above, the resolved
            // path must still sit inside the signatures folder, or it is refused.
            if (!candidate.StartsWith(folderPrefix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }
}
