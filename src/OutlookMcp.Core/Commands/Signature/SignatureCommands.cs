using OutlookMcp.Core.Models;

namespace OutlookMcp.Core.Commands.Signature;

/// <summary>
/// Signature discovery and reading. This is deliberately a plain filesystem command, not a COM one:
/// Outlook keeps signatures as files under <c>%APPDATA%\Microsoft\Signatures</c> and exposes no
/// signature object through the Outlook object model. Reading them here is read-only and complements
/// mail composition — a caller can fetch a signature's text and append it to a draft. Because there
/// is no COM involved, these methods do not go through <c>OutlookInteropRunner</c>.
/// </summary>
public class SignatureCommands : ISignatureCommands
{
    private readonly string _signaturesFolder;

    public SignatureCommands()
        : this(GetDefaultSignaturesFolder())
    {
    }

    /// <summary>
    /// Overload for tests: scan an explicit folder instead of the real
    /// <c>%APPDATA%\Microsoft\Signatures</c>.
    /// </summary>
    internal SignatureCommands(string signaturesFolder)
    {
        _signaturesFolder = signaturesFolder;
    }

    internal static string GetDefaultSignaturesFolder()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        return Path.Combine(appData, "Microsoft", "Signatures");
    }

    public OutlookSignatureListResult ListSignatures()
    {
        try
        {
            bool folderExists = Directory.Exists(_signaturesFolder);
            var signatures = SignatureFileScanner.Scan(_signaturesFolder);

            return new OutlookSignatureListResult
            {
                Success = true,
                Signatures = signatures,
                Count = signatures.Count,
                SignaturesFolderPath = _signaturesFolder,
                FolderExists = folderExists
            };
        }
        catch (Exception ex)
        {
            return new OutlookSignatureListResult
            {
                Success = false,
                SignaturesFolderPath = _signaturesFolder,
                ErrorMessage = $"Failed to list Outlook signatures: {ex.Message}"
            };
        }
    }

    public OutlookSignatureReadResult ReadSignature(string signatureName, string format = "text")
    {
        if (string.IsNullOrWhiteSpace(signatureName))
        {
            return new OutlookSignatureReadResult
            {
                Success = false,
                ErrorMessage = "signatureName is required."
            };
        }

        var normalizedFormat = (format ?? "text").Trim().ToLowerInvariant();
        normalizedFormat = normalizedFormat switch
        {
            "htm" or "html" => "html",
            "txt" or "text" or "plain" => "text",
            "rtf" => "rtf",
            _ => normalizedFormat
        };

        if (normalizedFormat is not ("html" or "text" or "rtf"))
        {
            return new OutlookSignatureReadResult
            {
                Success = false,
                Name = signatureName,
                ErrorMessage = $"Unknown format '{format}'. Use one of: text, html, rtf."
            };
        }

        try
        {
            if (!Directory.Exists(_signaturesFolder))
            {
                return new OutlookSignatureReadResult
                {
                    Success = false,
                    Name = signatureName,
                    ErrorMessage = $"The signatures folder does not exist ({_signaturesFolder}); "
                        + "no signatures are defined."
                };
            }

            var path = SignatureFileScanner.ResolveSignatureFile(_signaturesFolder, signatureName, normalizedFormat);
            if (path is null)
            {
                var available = SignatureFileScanner.Scan(_signaturesFolder)
                    .FirstOrDefault(s => string.Equals(s.Name, signatureName, StringComparison.OrdinalIgnoreCase));

                if (available is null)
                {
                    return new OutlookSignatureReadResult
                    {
                        Success = false,
                        Name = signatureName,
                        ErrorMessage = $"No signature named '{signatureName}' was found in {_signaturesFolder}."
                    };
                }

                var formats = new List<string>();
                if (available.HasText) formats.Add("text");
                if (available.HasHtml) formats.Add("html");
                if (available.HasRtf) formats.Add("rtf");

                return new OutlookSignatureReadResult
                {
                    Success = false,
                    Name = signatureName,
                    ErrorMessage = $"Signature '{signatureName}' has no {normalizedFormat} form. "
                        + $"Available formats: {string.Join(", ", formats)}."
                };
            }

            var content = File.ReadAllText(path);

            return new OutlookSignatureReadResult
            {
                Success = true,
                Name = signatureName,
                Format = normalizedFormat,
                Content = content,
                SourcePath = path
            };
        }
        catch (Exception ex)
        {
            return new OutlookSignatureReadResult
            {
                Success = false,
                Name = signatureName,
                ErrorMessage = $"Failed to read signature '{signatureName}': {ex.Message}"
            };
        }
    }
}
