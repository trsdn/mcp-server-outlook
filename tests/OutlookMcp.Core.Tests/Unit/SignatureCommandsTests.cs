using OutlookMcp.Core.Commands.Signature;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Tests for <see cref="SignatureCommands"/> driven against a temporary signatures folder via its
/// internal test constructor. Signatures are plain files, not COM objects — the command never
/// touches Outlook — so this is pure filesystem logic covered by the narrow Rule 30 / ADR-001
/// exception. Each test writes real files and asserts the command's result, so it fails if the
/// listing or reading logic is wrong.
/// </summary>
public class SignatureCommandsTests : IDisposable
{
    private readonly string _folder;

    public SignatureCommandsTests()
    {
        _folder = Path.Combine(
            AppContext.BaseDirectory,
            "sig-command-tests",
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
            // Best-effort cleanup.
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ListSignatures_ReturnsSignaturesAndFolderState()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.htm"), "<p>work</p>");
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "work");

        var result = new SignatureCommands(_folder).ListSignatures();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.FolderExists);
        Assert.Equal(1, result.Count);
        Assert.Equal(_folder, result.SignaturesFolderPath);
        var work = Assert.Single(result.Signatures);
        Assert.Equal("Work", work.Name);
        Assert.True(work.HasHtml);
        Assert.True(work.HasText);
    }

    [Fact]
    public void ListSignatures_MissingFolder_SucceedsWithEmptyList()
    {
        var missing = Path.Combine(_folder, "gone");

        var result = new SignatureCommands(missing).ListSignatures();

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.FolderExists);
        Assert.Empty(result.Signatures);
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void ReadSignature_ReturnsRequestedFormatContent()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "Best regards");

        var result = new SignatureCommands(_folder).ReadSignature("Work", "text");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("Work", result.Name);
        Assert.Equal("text", result.Format);
        Assert.Equal("Best regards", result.Content);
    }

    [Fact]
    public void ReadSignature_UnknownSignature_FailsWithoutSuccess()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "x");

        var result = new SignatureCommands(_folder).ReadSignature("Missing", "text");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Null(result.Content);
    }

    [Fact]
    public void ReadSignature_FormatNotPresent_ReportsAvailableFormats()
    {
        File.WriteAllText(Path.Combine(_folder, "Work.txt"), "x");

        var result = new SignatureCommands(_folder).ReadSignature("Work", "html");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("text", result.ErrorMessage);
    }

    [Fact]
    public void ReadSignature_BlankName_Fails()
    {
        var result = new SignatureCommands(_folder).ReadSignature("  ", "text");

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void ReadSignature_RootedName_DoesNotEscapeTheSignaturesFolder()
    {
        // Prove the command does not read a file outside the signatures folder via a rooted name
        // (Path.Combine would otherwise honour the rooted path and discard the base folder).
        var outsideDir = Path.Combine(
            AppContext.BaseDirectory, "sig-command-outside", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outsideDir);
        try
        {
            File.WriteAllText(Path.Combine(outsideDir, "secret.txt"), "top secret");
            var rootedName = Path.Combine(outsideDir, "secret");

            var result = new SignatureCommands(_folder).ReadSignature(rootedName, "text");

            Assert.False(result.Success);
            Assert.Null(result.Content);
            Assert.DoesNotContain("top secret", result.Content ?? string.Empty);
        }
        finally
        {
            Directory.Delete(outsideDir, recursive: true);
        }
    }
}
