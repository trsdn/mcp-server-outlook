using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.File;

/// <summary>
/// Legacy file management commands for inherited PowerPoint presentations.
/// These remain during migration but are not part of the active Outlook-first seed.
/// </summary>
[ServiceCategory("file")]
[NoSession]
public interface IFileCommands
{
    /// <summary>
    /// Validate a legacy PowerPoint file and return metadata (size, slide count, macro status).
    /// </summary>
    /// <param name="filePath">Path to the .pptx or .pptm file</param>
    [ServiceAction("test")]
    FileValidationInfo Test(string filePath);
}
