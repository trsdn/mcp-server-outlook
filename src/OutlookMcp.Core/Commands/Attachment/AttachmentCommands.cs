using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Attachment;

public class AttachmentCommands : IAttachmentCommands
{
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AttachmentMutationResult Add(
        string filePath,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = false)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return new AttachmentMutationResult
            {
                Success = false,
                ErrorMessage = "filePath is required for attachment.add."
            };
        }

        if (!System.IO.File.Exists(filePath))
        {
            return new AttachmentMutationResult
            {
                Success = false,
                ErrorMessage = $"Attachment file was not found: {filePath}"
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookAttachmentAdd",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                Outlook.Attachments? attachments = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        mailEntryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new AttachmentMutationResult
                        {
                            Success = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(mailEntryId)
                                ? "No active Outlook mail item is currently selected or open."
                                : "The requested Outlook mail item could not be resolved."
                        };
                    }

                    if (SafeGetBool(() => mail.Sent))
                    {
                        return new AttachmentMutationResult
                        {
                            Success = false,
                            ErrorMessage = "Attachments can only be added to Outlook draft items."
                        };
                    }

                    attachments = mail.Attachments;
                    _ = attachments.Add(filePath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    mail.Save();

                    return new AttachmentMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        AttachmentCount = SafeGetInt(() => attachments.Count),
                        FileName = Path.GetFileName(filePath),
                        Message = "Added Outlook attachment to draft."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref attachments);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new AttachmentMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to add the Outlook attachment: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AttachmentListResult List(
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookAttachmentList",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                Outlook.Attachments? attachments = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        mailEntryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new AttachmentListResult
                        {
                            Success = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(mailEntryId)
                                ? "No active Outlook mail item is currently selected or open."
                                : "The requested Outlook mail item could not be resolved."
                        };
                    }

                    attachments = mail.Attachments;
                    var result = new AttachmentListResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        AttachmentCount = SafeGetInt(() => attachments.Count)
                    };

                    for (int index = 1; index <= result.AttachmentCount; index++)
                    {
                        Outlook.Attachment? attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            result.Attachments.Add(new AttachmentInfo
                            {
                                Index = index,
                                FileName = SafeGet(() => attachment.FileName) ?? $"attachment-{index}",
                                SizeBytes = SafeGetInt(() => attachment.Size),
                                ContentId = SafeGet(() => attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F") as string),
                                DisplayName = SafeGet(() => attachment.DisplayName),
                                Type = SafeGet(() => attachment.Type.ToString()),
                                Hidden = SafeGetBool(() => attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B") is true)
                            });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref attachment);
                        }
                    }

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref attachments);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new AttachmentListResult
            {
                Success = false,
                ErrorMessage = $"Failed to inspect Outlook attachments: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AttachmentSaveResult Save(
        string destinationDirectory,
        int attachmentIndex = 0,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = true,
        bool overwrite = false)
    {
        if (string.IsNullOrWhiteSpace(destinationDirectory))
        {
            return new AttachmentSaveResult
            {
                Success = false,
                ErrorMessage = "destinationDirectory is required for attachment.save."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookAttachmentSave",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                Outlook.Attachments? attachments = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        mailEntryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new AttachmentSaveResult
                        {
                            Success = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(mailEntryId)
                                ? "No active Outlook mail item is currently selected or open."
                                : "The requested Outlook mail item could not be resolved."
                        };
                    }

                    Directory.CreateDirectory(destinationDirectory);

                    attachments = mail.Attachments;
                    int count = SafeGetInt(() => attachments.Count);
                    if (count == 0)
                    {
                        return new AttachmentSaveResult
                        {
                            Success = true,
                            EntryId = SafeGet(() => mail.EntryID),
                            StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                            Subject = SafeGet(() => mail.Subject),
                            SavedCount = 0,
                            Message = "The selected Outlook mail item has no attachments."
                        };
                    }

                    var indexes = attachmentIndex > 0
                        ? [attachmentIndex]
                        : Enumerable.Range(1, count).ToArray();

                    if (indexes.Any(index => index < 1 || index > count))
                    {
                        return new AttachmentSaveResult
                        {
                            Success = false,
                            ErrorMessage = $"attachmentIndex must be between 1 and {count}, or 0 to save all attachments."
                        };
                    }

                    var plannedSaves = new List<(int Index, string FilePath)>(indexes.Length);
                    foreach (int index in indexes)
                    {
                        Outlook.Attachment? attachment = null;
                        try
                        {
                            attachment = attachments[index];
                            string fileName = SanitizeFileName(SafeGet(() => attachment.FileName) ?? $"attachment-{index}.bin");
                            string filePath = Path.Combine(destinationDirectory, fileName);
                            if (!overwrite && System.IO.File.Exists(filePath))
                            {
                                return new AttachmentSaveResult
                                {
                                    Success = false,
                                    ErrorMessage = $"Attachment destination already exists: {filePath}"
                                };
                            }

                            plannedSaves.Add((index, filePath));
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref attachment);
                        }
                    }

                    var savedFiles = new List<string>(plannedSaves.Count);
                    foreach (var plannedSave in plannedSaves)
                    {
                        Outlook.Attachment? attachment = null;
                        try
                        {
                            attachment = attachments[plannedSave.Index];
                            attachment.SaveAsFile(plannedSave.FilePath);
                            savedFiles.Add(plannedSave.FilePath);
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref attachment);
                        }
                    }

                    return new AttachmentSaveResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        SavedCount = savedFiles.Count,
                        SavedFiles = savedFiles,
                        Message = savedFiles.Count == 1
                            ? "Saved 1 Outlook attachment."
                            : $"Saved {savedFiles.Count} Outlook attachments."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref attachments);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new AttachmentSaveResult
            {
                Success = false,
                ErrorMessage = $"Failed to save Outlook attachments: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public AttachmentMutationResult Remove(
        int attachmentIndex,
        string? mailEntryId = null,
        string? storeId = null,
        bool useActiveMail = false,
        bool confirm = false)
    {
        if (attachmentIndex < 1)
        {
            return new AttachmentMutationResult
            {
                Success = false,
                ErrorMessage = "attachmentIndex must be 1 or greater for attachment.remove."
            };
        }

        // Checked after the index, and before Outlook is reached at all: an index of 0 is a caller
        // mistake worth naming on its own, and answering it with "pass confirm=true" would send the
        // caller to fix the wrong thing. See #9.
        if (!confirm)
        {
            return new AttachmentMutationResult
            {
                Success = false,
                ErrorMessage = ConfirmationGate.AttachmentRemove()
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookAttachmentRemove",
            (application, session) =>
            {
                Outlook.MailItem? mail = null;
                Outlook.Inspector? inspector = null;
                Outlook.Explorer? explorer = null;
                Outlook.Selection? selection = null;
                Outlook.Attachments? attachments = null;
                Outlook.Attachment? attachment = null;
                object? currentItem = null;
                object? selectedItem = null;
                object? resolvedItem = null;

                try
                {
                    mail = OutlookInteropRunner.ResolveMailItem(
                        application,
                        session,
                        mailEntryId,
                        storeId,
                        useActiveMail,
                        out inspector,
                        out explorer,
                        out selection,
                        out currentItem,
                        out selectedItem,
                        out resolvedItem);

                    if (mail == null)
                    {
                        return new AttachmentMutationResult
                        {
                            Success = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(mailEntryId)
                                ? "No active Outlook mail item is currently selected or open."
                                : "The requested Outlook mail item could not be resolved."
                        };
                    }

                    if (SafeGetBool(() => mail.Sent))
                    {
                        return new AttachmentMutationResult
                        {
                            Success = false,
                            ErrorMessage = "Attachments can only be removed from Outlook draft items."
                        };
                    }

                    attachments = mail.Attachments;
                    int count = SafeGetInt(() => attachments.Count);
                    if (attachmentIndex > count)
                    {
                        return new AttachmentMutationResult
                        {
                            Success = false,
                            ErrorMessage = $"attachmentIndex must be between 1 and {count}."
                        };
                    }

                    attachment = attachments[attachmentIndex];
                    string fileName = SafeGet(() => attachment.FileName) ?? $"attachment-{attachmentIndex}";
                    attachment.Delete();
                    mail.Save();

                    return new AttachmentMutationResult
                    {
                        Success = true,
                        EntryId = SafeGet(() => mail.EntryID),
                        StoreId = SafeGet(() => mail.Parent is Outlook.MAPIFolder folder ? folder.StoreID : null),
                        Subject = SafeGet(() => mail.Subject),
                        AttachmentCount = SafeGetInt(() => attachments.Count),
                        FileName = fileName,
                        Message = "Removed Outlook attachment from draft."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref attachment);
                    OutlookInteropRunner.ReleaseComObject(ref attachments);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedItem);
                    OutlookInteropRunner.ReleaseComObject(ref selectedItem);
                    OutlookInteropRunner.ReleaseComObject(ref currentItem);
                    OutlookInteropRunner.ReleaseComObject(ref selection);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                    OutlookInteropRunner.ReleaseComObject(ref inspector);
                    OutlookInteropRunner.ReleaseComObject(ref mail);
                }
            },
            ex => new AttachmentMutationResult
            {
                Success = false,
                ErrorMessage = $"Failed to remove the Outlook attachment: {ex.Message}"
            });
    }

    private static string SanitizeFileName(string fileName)
    {
        char[] invalidChars = Path.GetInvalidFileNameChars();
        return string.Concat(fileName.Select(ch => invalidChars.Contains(ch) ? '_' : ch));
    }

    private static string? SafeGet(Func<string?> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static bool SafeGetBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return false;
        }
    }

    private static int SafeGetInt(Func<int> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return 0;
        }
    }
}
