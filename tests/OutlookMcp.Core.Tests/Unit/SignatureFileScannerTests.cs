using OutlookMcp.Core.Commands.Signature;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Tests for <see cref="SignatureFileScanner"/>, which groups the files Outlook keeps under
/// <c>%APPDATA%\Microsoft\Signatures</c> into one signature per base name.
///
/// This is pure filesystem logic with zero COM dependency — no Outlook type appears anywhere in
/// <see cref="SignatureFileScanner"/>'s signature — so it falls under the narrow exception Rule 30
/// and ADR-001 carve out for genuinely pure logic. Each test writes real files to a unique temporary
/// directory and asserts the grouping, so it fails if the grouping logic is wrong (not merely if
/// .NET is broken). The real <c>%APPDATA%</c> folder is exercised separately by the integration
/// suite, which skips when the profile has no signatures.
/// </summary>
public class SignatureFileScannerTests : IDisposable
{
    private readonly string _folder;

    public SignatureFileScannerTests()
    {
        _folder = Path.Combine(
            AppContext.BaseDirectory,
            "sig-scanner-tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_folder);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_folder))
            {
                Directory.Delete(_folder, recursive: true);
            }
        }
        catch
        {
            // Best-effort cleanup of a temp folder; a leftover directory must not fail the suite.
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Scan_MissingFolder_ReturnsEmpty()
    {
        var missing = Path.Combine(_folder, "does-not-exist");

        var result = SignatureFileScanner.Scan(missing);

        Assert.Empty(result);
    }

    [Fact]
    public void Scan_GroupsThreeFilesOfOneSignatureIntoOneEntry()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.htm"), "<p>work</p>");
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "work");
        File.WriteAllText(Path.Combine(_folder, "Work.rtf"), "{\\rtf1 work}");

        var result = SignatureFileScanner.Scan(_folder);

        var work = Assert.Single(result);
        Assert.Equal("Work", work.Name);
        Assert.True(work.HasHtml);
        Assert.True(work.HasText);
        Assert.True(work.HasRtf);
    }

    [Fact]
    public void Scan_ReportsMissingFormatsPerSignature()
    {
        File.WriteAllText(Path.Combine(_folder, "PlainOnly.txt"), "hi");

        var result = SignatureFileScanner.Scan(_folder);

        var sig = Assert.Single(result);
        Assert.Equal("PlainOnly", sig.Name);
        Assert.True(sig.HasText);
        Assert.False(sig.HasHtml);
        Assert.False(sig.HasRtf);
    }

    [Fact]
    public void Scan_ReturnsMultipleSignaturesSortedByName()
    {
        File.WriteAllText(Path.Combine(_folder, "Zeta.htm"), "z");
        File.WriteAllText(Path.Combine(_folder, "alpha.txt"), "a");

        var result = SignatureFileScanner.Scan(_folder);

        Assert.Equal(2, result.Count);
        Assert.Equal("alpha", result[0].Name);
        Assert.Equal("Zeta", result[1].Name);
    }

    [Fact]
    public void Scan_IgnoresUnrelatedFiles()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "work");
        File.WriteAllText(Path.Combine(_folder, "notes.docx"), "x");
        Directory.CreateDirectory(Path.Combine(_folder, "Work_files"));

        var result = SignatureFileScanner.Scan(_folder);

        var work = Assert.Single(result);
        Assert.Equal("Work", work.Name);
    }

    [Fact]
    public void ResolveSignatureFile_FindsRequestedFormat()
    {
        var htmPath = Path.Combine(_folder, "Work.htm");
        File.WriteAllText(htmPath, "<p>work</p>");
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "work");

        var resolved = SignatureFileScanner.ResolveSignatureFile(_folder, "Work", "html");

        Assert.Equal(htmPath, resolved);
    }

    [Fact]
    public void ResolveSignatureFile_ReturnsNullWhenFormatMissing()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "work");

        var resolved = SignatureFileScanner.ResolveSignatureFile(_folder, "Work", "rtf");

        Assert.Null(resolved);
    }
}
