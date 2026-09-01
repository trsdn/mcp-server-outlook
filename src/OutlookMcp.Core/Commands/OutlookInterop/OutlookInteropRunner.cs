using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.ComInterop;
using OutlookMcp.ComInterop.Session;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.OutlookInterop;

internal static class OutlookInteropRunner
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    /// <summary>
    /// Dispatches an Outlook COM operation onto the single process-wide
    /// <see cref="OutlookDispatcher"/> STA thread (see #20 / ADR-002). Resolves the shared
    /// <c>Outlook.Application</c>/<c>Outlook.NameSpace</c> fresh for every call (Outlook itself is
    /// a single already-running singleton; resolving it does not create work), runs
    /// <paramref name="action"/>, and releases the per-call <c>NameSpace</c> — but never
    /// final-releases the shared <c>Application</c> RCW, per #19.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static TResult Execute<TResult>(
        string operationName,
        Func<Outlook.Application, Outlook.NameSpace, TResult> action,
        Func<Exception, TResult> onException)
    {
        return OutlookDispatcher.Shared.Execute(operationName, () =>
        {
            Outlook.Application? application = null;
            Outlook.NameSpace? session = null;

            try
            {
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    OutlookFlavor flavor = OutlookInstallationDetector.DetectFlavor();
                    return onException(new InvalidOperationException(
                        OutlookInstallationDetector.BuildUnavailableMessage(flavor)));
                }

                if (!TryGetRunningApplication(out application))
                {
                    // Deliberately do NOT fall back to Activator.CreateInstance(outlookType) here
                    // (see #30). A freshly created Outlook.Application is not the user's trusted,
                    // already-running session, so it is *more* likely to trigger the Outlook Object
                    // Model Guard (OMG) than a session obtained via GetActiveObject, and it conflicts
                    // with Outlook's single-instance-per-user model. Fail with actionable guidance
                    // instead, distinguishing "not running" from "running at a different integrity
                    // level than this (possibly elevated) process" -- GetActiveObject fails with
                    // MK_E_UNAVAILABLE in the latter case and TryGetRunningApplication surfaces that
                    // as a plain "not running" result, so we check elevation separately to give a
                    // more specific message.
                    bool elevated = OutlookInstallationDetector.IsCurrentProcessElevated();
                    string message = elevated
                        ? "Outlook.Application COM ProgID is registered, but no running instance could be reached. " +
                          "This process is running elevated (as Administrator); if Outlook is running unelevated, " +
                          "COM's GetActiveObject cannot see across integrity levels. Run Outlook and this server at " +
                          "the same elevation level, or start classic Outlook for Windows and try again."
                        : "Classic Outlook for Windows is installed but does not appear to be running. " +
                          "Start Outlook and try again.";
                    return onException(new InvalidOperationException(message));
                }

                session = application.GetNamespace("MAPI");

                return action(application, session);
            }
            catch (COMException ex) when (IsObjectModelGuardDenial(ex))
            {
                // The Outlook Object Model Guard (OMG) raises a modal "an application is trying to
                // access e-mail addresses" / "trying to send e-mail" dialog for out-of-process,
                // untrusted callers touching protected members (Recipients, SenderEmailAddress,
                // MailItem.Send(), AddressEntry.Address, etc.). If the user does not respond, Outlook
                // eventually aborts the call with E_ABORT / MK_E_UNAVAILABLE-shaped COMExceptions.
                // Surface this distinctly (Rule 22: never swallow security denials into an ambiguous
                // null/false/generic-error result) rather than reporting a plain COM failure. See #30.
                return onException(new InvalidOperationException(
                    "Outlook blocked this operation with a security prompt (Object Model Guard). " +
                    "This typically happens when reading a sensitive property (e.g. SenderEmailAddress, " +
                    "Recipients) or sending mail from an untrusted, out-of-process caller. If a dialog is " +
                    "visible in Outlook, a person must approve or dismiss it; this server cannot answer it " +
                    "automatically. See docs/ADR-002-OUTLOOK-COM-EXECUTION-MODEL.md for the documented OMG " +
                    $"posture and mitigations. Original error: {ex.Message}", ex));
            }
            catch (COMException ex)
            {
                return onException(ex);
            }
            catch (Exception ex)
            {
                return onException(ex);
            }
            finally
            {
                ReleaseComObject(ref session);
                // Do NOT final-release `application`: it is the user's already-running,
                // shared Outlook.Application instance (obtained via GetActiveObject and
                // cached per-process by the RCW table). Final-releasing it zeroes the RCW
                // refcount for every holder in the process, not just this call, and can
                // invalidate a later operation's reference to the same object. See #19.
                ReleaseSharedComObject(ref application);
            }
        }, ComInteropConstants.DefaultOperationTimeout);
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

    /// <summary>
    /// Identifies COMExceptions consistent with an Outlook Object Model Guard (OMG) denial: a
    /// dismissed or unanswered "an application is trying to..." security prompt. OMG-triggered
    /// failures surface as <c>E_ABORT</c> (0x80004004, "Operation aborted") from the Outlook
    /// automation surface, since OMG aborts the call rather than returning a normal error result.
    /// This is a heuristic (Outlook does not expose a dedicated "OMG denied" HRESULT) but is
    /// specific enough in practice to distinguish from generic COM failures. See #30.
    /// </summary>
    internal static bool IsObjectModelGuardDenial(COMException ex)
    {
        const int E_ABORT = unchecked((int)0x80004004);
        return ex.HResult == E_ABORT;
    }

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

    /// <summary>
    /// Releases a COM reference we do NOT own the lifetime of (e.g. the shared, already-running
    /// Outlook.Application returned by GetActiveObject). Uses a plain ref-count decrement instead
    /// of FinalReleaseComObject so other holders of the same cached RCW are unaffected. See #19.
    /// </summary>
    internal static void ReleaseSharedComObject<T>(ref T? value) where T : class
    {
        if (value != null && Marshal.IsComObject(value))
        {
            _ = Marshal.ReleaseComObject(value);
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
