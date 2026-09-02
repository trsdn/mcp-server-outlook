using Xunit;

namespace OutlookMcp.SkillGeneration.Tests;

/// <summary>
/// Validates the active Outlook migration skill surfaces.
/// Validates the active Outlook skill surfaces: the repo's skill story is Outlook-first.
/// </summary>
public class SkillMdQualityTests
{
    private static readonly string SkillsFolder = Path.Combine(
        AppContext.BaseDirectory, "skills");

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void OutlookCliSkill_Exists()
    {
        var skillPath = Path.Combine(SkillsFolder, "outlook-cli", "SKILL.md");
        Assert.True(File.Exists(skillPath), $"Outlook CLI SKILL.md should exist at {skillPath}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void OutlookMcpSkill_Exists()
    {
        var skillPath = Path.Combine(SkillsFolder, "outlook-mcp", "SKILL.md");
        Assert.True(File.Exists(skillPath), $"Outlook MCP SKILL.md should exist at {skillPath}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void OutlookCliSkill_DescribesCurrentSeed()
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, "outlook-cli", "SKILL.md"));

        Assert.Contains("application.get-status", content);
        Assert.Contains("mail.list", content);
        Assert.Contains("mail.search", content);
        Assert.Contains("mail.send", content);
        Assert.Contains("attachment.list", content);
        Assert.Contains("attachment.save", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void OutlookMcpSkill_DescribesCurrentSeed()
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, "outlook-mcp", "SKILL.md"));

        Assert.Contains("folder.list-default", content);
        Assert.Contains("mail.read-active", content);
        Assert.Contains("mail.create-draft", content);
        Assert.Contains("mail.reply-all", content);
        Assert.Contains("attachment.save", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void OutlookMcpSkill_PrefersSafeDraftWorkflow()
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, "outlook-mcp", "SKILL.md"));

        Assert.Contains("Prefer draft-producing actions", content);
        Assert.Contains("mail.send", content);
        Assert.Contains("explicit final action", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void OutlookCliSkill_DocumentsTransitionalNaming()
    {
        var content = File.ReadAllText(Path.Combine(SkillsFolder, "outlook-cli", "SKILL.md"));

        Assert.Contains("outlookcli", content);
        Assert.Contains("OutlookMcp.CLI", content);
        Assert.Contains("migration", content, StringComparison.OrdinalIgnoreCase);
    }
}
