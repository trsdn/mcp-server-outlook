using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PptMcp.ComInterop;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PptMcp.Core.Commands.OutlookInterop;

internal static class OutlookInteropRunner
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static TResult Execute<TResult>(
        string operationName,
        Func<Outlook.Application, Outlook.NameSpace, TResult> action,
        Func<Exception, TResult> onException)
    {
        var completion = new TaskCompletionSource<TResult>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            Outlook.Application? application = null;
            Outlook.NameSpace? session = null;

            try
            {
                OleMessageFilter.Register();

                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    completion.TrySetResult(onException(new InvalidOperationException("Microsoft Outlook is not installed on this system.")));
                    return;
                }

                application = GetOrCreateApplication(outlookType);
                session = application.GetNamespace("MAPI");

                completion.TrySetResult(action(application, session));
            }
            catch (COMException ex)
            {
                completion.TrySetResult(onException(ex));
            }
            catch (Exception ex)
            {
                completion.TrySetResult(onException(ex));
            }
            finally
            {
                ReleaseComObject(ref session);
                ReleaseComObject(ref application);
                OleMessageFilter.Revoke();
            }
        })
        {
            IsBackground = true,
            Name = operationName
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        if (!completion.Task.Wait(ComInteropConstants.DefaultOperationTimeout))
        {
            throw new TimeoutException($"{operationName} timed out while waiting for Outlook.");
        }

        _ = thread.Join(TimeSpan.FromSeconds(10));
        return completion.Task.GetAwaiter().GetResult();
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static bool TryGetRunningApplication([NotNullWhen(true)] out Outlook.Application? application)
    {
        application = null;

        const string progId = "Outlook.Application";
        if (CLSIDFromProgID(progId, out Guid clsid) != 0)
        {
            return false;
        }

        if (GetActiveObject(ref clsid, IntPtr.Zero, out object? activeObject) != 0 || activeObject == null)
        {
            return false;
        }

        application = activeObject as Outlook.Application;
        if (application == null && Marshal.IsComObject(activeObject))
        {
            _ = Marshal.FinalReleaseComObject(activeObject);
        }

        return application != null;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static Outlook.Application GetOrCreateApplication(Type outlookType)
    {
        if (TryGetRunningApplication(out Outlook.Application? runningApplication))
        {
            return runningApplication;
        }

#pragma warning disable IL2072
        return (Outlook.Application)Activator.CreateInstance(outlookType)!;
#pragma warning restore IL2072
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static Outlook.MAPIFolder? ResolveFolder(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? folder,
        IReadOnlyDictionary<string, Outlook.OlDefaultFolders> aliases,
        ref Outlook.Explorer? explorer)
    {
        if (string.IsNullOrWhiteSpace(folder) || string.Equals(folder, "current", StringComparison.OrdinalIgnoreCase))
        {
            explorer ??= application.ActiveExplorer();
            if (explorer?.CurrentFolder is Outlook.MAPIFolder currentFolder)
            {
                return currentFolder;
            }

            return aliases.TryGetValue("inbox", out var inboxFolder)
                ? session.GetDefaultFolder(inboxFolder)
                : null;
        }

        if (aliases.TryGetValue(folder, out var defaultFolder))
        {
            return session.GetDefaultFolder(defaultFolder);
        }

        return FindFolderByPath(session, folder);
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? FindFolderByPath(Outlook.NameSpace session, string folderPath)
    {
        Outlook.Folders? rootFolders = null;
        try
        {
            string normalizedPath = NormalizeFolderIdentifier(folderPath);
            bool simpleName = !normalizedPath.Contains('\\');
            rootFolders = session.Folders;

            int count = rootFolders.Count;
            for (int index = 1; index <= count; index++)
            {
                Outlook.MAPIFolder? rootFolder = null;
                try
                {
                    rootFolder = rootFolders[index];
                    var found = FindFolderRecursive(rootFolder, normalizedPath, simpleName);
                    if (found != null)
                    {
                        if (ReferenceEquals(found, rootFolder))
                        {
                            rootFolder = null;
                        }

                        return found;
                    }
                }
                finally
                {
                    ReleaseComObject(ref rootFolder);
                }
            }

            return null;
        }
        finally
        {
            ReleaseComObject(ref rootFolders);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? FindFolderRecursive(
        Outlook.MAPIFolder folder,
        string normalizedPath,
        bool simpleName)
    {
        if (MatchesFolder(folder, normalizedPath, simpleName))
        {
            return folder;
        }

        Outlook.Folders? childFolders = null;
        try
        {
            childFolders = folder.Folders;
            int count = childFolders.Count;
            for (int index = 1; index <= count; index++)
            {
                Outlook.MAPIFolder? childFolder = null;
                try
                {
                    childFolder = childFolders[index];
                    var found = FindFolderRecursive(childFolder, normalizedPath, simpleName);
                    if (found != null)
                    {
                        if (ReferenceEquals(found, childFolder))
                        {
                            childFolder = null;
                        }

                        return found;
                    }
                }
                finally
                {
                    ReleaseComObject(ref childFolder);
                }
            }

            return null;
        }
        finally
        {
            ReleaseComObject(ref childFolders);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool MatchesFolder(Outlook.MAPIFolder folder, string normalizedPath, bool simpleName)
    {
        try
        {
            string folderName = folder.Name ?? string.Empty;
            if (simpleName && folderName.Equals(normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string? folderPath = GetFolderPath(folder);
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return false;
            }

            string normalizedFolderPath = NormalizeFolderIdentifier(folderPath);
            return normalizedFolderPath.Equals(normalizedPath, StringComparison.OrdinalIgnoreCase)
                   || normalizedFolderPath.EndsWith("\\" + normalizedPath.TrimStart('\\'), StringComparison.OrdinalIgnoreCase);
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static string NormalizeFolderIdentifier(string value)
    {
        string normalized = value
            .Replace('/', '\\')
            .Trim();

        while (normalized.StartsWith(@"\\", StringComparison.Ordinal))
        {
            normalized = normalized[1..];
        }

        return normalized.Trim('\\');
    }

    internal static string NormalizeBodyPreview(string? body, int maxLength = 500)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return string.Empty;
        }

        string normalized = body
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim();

        if (normalized.Length <= maxLength)
        {
            return normalized;
        }

        return normalized[..maxLength].TrimEnd() + "...";
    }

    internal static string? GetFolderPath(Outlook.MAPIFolder? folder)
    {
        if (folder == null)
        {
            return null;
        }

        try
        {
            return folder.FolderPath;
        }
        catch (COMException)
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static Outlook.MailItem? ResolveMailItem(
        Outlook.Application application,
        Outlook.NameSpace session,
        string? entryId,
        string? storeId,
        bool useActiveMail,
        out Outlook.Inspector? inspector,
        out Outlook.Explorer? explorer,
        out Outlook.Selection? selection,
        out object? currentItem,
        out object? selectedItem,
        out object? resolvedItem)
    {
        inspector = null;
        explorer = null;
        selection = null;
        currentItem = null;
        selectedItem = null;
        resolvedItem = null;

        if (!string.IsNullOrWhiteSpace(entryId))
        {
            resolvedItem = session.GetItemFromID(
                entryId,
                string.IsNullOrWhiteSpace(storeId) ? Type.Missing : storeId);
            return resolvedItem as Outlook.MailItem;
        }

        if (!useActiveMail)
        {
            return null;
        }

        inspector = application.ActiveInspector();
        if (inspector != null)
        {
            currentItem = inspector.CurrentItem;
            if (currentItem is Outlook.MailItem currentMail)
            {
                return currentMail;
            }
        }

        explorer = application.ActiveExplorer();
        if (explorer != null)
        {
            selection = explorer.Selection;
            if (selection != null && selection.Count > 0)
            {
                selectedItem = selection[1];
                if (selectedItem is Outlook.MailItem selectedMail)
                {
                    return selectedMail;
                }
            }
        }

        return null;
    }

    internal static void ReleaseComObject<T>(ref T? value) where T : class
    {
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }

        value = null;
    }

    internal static void ReleaseComObject(ref object? value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }

        value = null;
    }
}
