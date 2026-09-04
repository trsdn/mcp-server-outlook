using OutlookMcp.Core.Commands.Signature;
using Xunit;
using Xunit.Abstractions;

namespace OutlookMcp.Core.Tests.Integration.Outlook;

/// <summary>
/// Signature listing against the real <c>%APPDATA%\Microsoft\Signatures</c> folder (#15).
///
/// <para>
/// Signatures are files, not COM objects, so this needs no running Outlook — but it does depend on
/// the owner's real profile, which may have no signatures at all. That is a legitimate state, so the
/// content assertions skip when the folder is empty rather than passing vacuously. The grouping and
/// reading logic itself is covered deterministically by the unit tests against a temporary folder.
/// </para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Signature")]
[Collection("Sequential")]
public class OutlookSignatureTests(ITestOutputHelper output)
{
    [Fact]
    public void ListSignatures_AlwaysSucceedsAndReportsTheFolder()
    {
        var result = new SignatureCommands().ListSignatures();

        // Listing is filesystem-only and must never fail: an absent folder simply means no signatures.
        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.SignaturesFolderPath));
        Assert.Equal(result.Count, result.Signatures.Count);

        output.WriteLine(
            $"folder={result.SignaturesFolderPath} exists={result.FolderExists} count={result.Count}");
        foreach (var signature in result.Signatures)
        {
            output.WriteLine(
                $"  {signature.Name} (html={signature.HasHtml} text={signature.HasText} rtf={signature.HasRtf})");
        }
    }

    [SkippableFact]
    public void ReadSignature_ReadsARealSignatureWhenOneExists()
    {
        var list = new SignatureCommands().ListSignatures();
        Assert.True(list.Success, list.ErrorMessage);
        Skip.If(list.Count == 0,
            "This profile has no signatures defined, so there is nothing to read. "
            + "The read logic is covered by the unit tests.");

        var target = list.Signatures.First();
        var format = target.HasText ? "text" : target.HasHtml ? "html" : "rtf";

        var result = new SignatureCommands().ReadSignature(target.Name, format);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(target.Name, result.Name);
        Assert.Equal(format, result.Format);
        Assert.NotNull(result.Content);
        output.WriteLine($"read '{target.Name}' as {format}: {result.Content!.Length} chars");
    }
}
