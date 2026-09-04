using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.ComInterop;
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
                    // Reaching here means Outlook could not be resolved by either route: it is
                    // absent from the Running Object Table *and* no OUTLOOK.EXE is running in this
                    // session (see TryGetRunningApplication, #90). Nothing here starts Outlook - a
                    // COM server this process spawned would put an Office window on the user's
                    // desktop unbidden, which this server must never do.
                    bool elevated = OutlookInstallationDetector.IsCurrentProcessElevated();
                    string message = elevated
                        ? "Outlook.Application COM ProgID is registered, but no running instance could be reached. " +
                          "This process is running elevated (as Administrator); if Outlook is running unelevated, " +
                          "COM cannot see across integrity levels. Run Outlook and this server at " +
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

    /// <summary>
    /// Resolves the Outlook instance the user already has open.
    ///
    /// <para>
    /// The Running Object Table is tried first, but its absence is not evidence: Outlook does not
    /// reliably register itself there. Observed with Outlook running, mailbox loaded and fully
    /// usable, <c>GetActiveObject</c> still returning <c>MK_E_UNAVAILABLE</c> - which the caller
    /// then reported as "Outlook does not appear to be running" (#90). So when the ROT lookup comes
    /// back empty we fall through to <c>CoCreateInstance</c>, which for Outlook - a single-instance
    /// COM server - returns a reference to the <em>same</em> running instance rather than a second
    /// one.
    /// </para>
    ///
    /// <para>
    /// That fallback is gated on an <c>OUTLOOK.EXE</c> process already existing in this session,
    /// and the gate is the whole point: <c>CoCreateInstance</c> with Outlook closed would
    /// <em>start</em> Outlook, and this server must never put an Office window on someone's desktop
    /// unbidden. With the gate, the fallback can only ever attach to something already running.
    /// </para>
    ///
    /// <para>
    /// This supersedes the earlier reasoning (#30) that any fallback to <c>CreateInstance</c> risks
    /// a "less trusted" Application and extra Object Model Guard prompts. There is no separate
    /// object to be less trusted - it is the user's own running session either way, reached from the
    /// same caller process, so the OMG posture is identical. Only the launch risk was real, and the
    /// gate removes it.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static bool TryGetRunningApplication([NotNullWhen(true)] out Outlook.Application? application)
    {
        application = null;

        const string progId = "Outlook.Application";
        if (CLSIDFromProgID(progId, out Guid clsid) != 0)
        {
            return false;
        }

        if (GetActiveObject(ref clsid, IntPtr.Zero, out object? activeObject) == 0 && activeObject != null)
        {
            application = activeObject as Outlook.Application;
            if (application == null && Marshal.IsComObject(activeObject))
            {
                _ = Marshal.FinalReleaseComObject(activeObject);
            }

            if (application != null)
            {
                return true;
            }
        }

        return TryAttachToRunningOutlookProcess(out application);
    }

    /// <summary>
    /// Attaches to an Outlook that is running but absent from the Running Object Table. Returns
    /// false without touching COM when no Outlook process exists in this session, so this can never
    /// launch Outlook. See <see cref="TryGetRunningApplication"/> and #90.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool TryAttachToRunningOutlookProcess([NotNullWhen(true)] out Outlook.Application? application)
    {
        application = null;

        if (!IsOutlookRunningInCurrentSession())
        {
            return false;
        }

        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType == null)
        {
            return false;
        }

        try
        {
            object? instance = Activator.CreateInstance(outlookType);
            application = instance as Outlook.Application;

            if (application == null && instance != null && Marshal.IsComObject(instance))
            {
                _ = Marshal.FinalReleaseComObject(instance);
            }

            return application != null;
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>
    /// Whether Outlook is running <em>in this Windows session</em>. Session-scoped deliberately: a
    /// COM attach cannot reach another session's Outlook, so counting one would turn the launch
    /// guard above into no guard at all.
    /// </summary>
    private static bool IsOutlookRunningInCurrentSession()
    {
        try
        {
            using Process current = Process.GetCurrentProcess();
            int sessionId = current.SessionId;

            foreach (Process process in Process.GetProcessesByName("OUTLOOK"))
            {
                try
                {
                    if (process.SessionId == sessionId)
                    {
                        return true;
                    }
                }
                catch (InvalidOperationException)
                {
                    // Exited between enumeration and inspection.
                }
                finally
                {
                    process.Dispose();
                }
            }
        }
        catch (InvalidOperationException)
        {
            return false;
        }

        return false;
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
    ///
    /// <para>
    /// Deliberately <em>not</em> widened to <c>MAPI_E_NOT_SUPPORTED</c> (0x80040102). Microsoft
    /// documents the guard returning that HRESULT for a protected member refused outright, but it
    /// is also the ordinary MAPI "this provider, store or property type does not support that"
    /// error - <c>PropertyAccessor</c> returns it for a <c>PT_OBJECT</c> property, and a computed
    /// user property returns it when it cannot be evaluated. Treating it as a denial here would
    /// make every such failure tell the caller to look for a security dialog that does not exist.
    /// Code that needs the distinction should triage the HRESULT itself and report the ambiguity
    /// rather than resolving it - see <c>PropertyCommands.ReadProperty</c>.
    /// </para>
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

            if (!simpleName)
            {
                var direct = WalkFolderPath(session, normalizedPath);
                if (direct != null)
                {
                    return direct;
                }
            }

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

    /// <summary>
    /// Resolves a full folder path one segment at a time, using Outlook's <c>Folders["name"]</c>
    /// indexer rather than enumerating the tree.
    ///
    /// <para>
    /// <b>This exists because enumerating goes stale.</b> A folder renamed in this process is not
    /// found by the recursive walk afterwards - the enumerated child objects still report the old
    /// name, and no amount of retrying refreshes them, so a path the tool had just handed back
    /// resolved to nothing. Verified intermittently, roughly one run in three. The name indexer asks
    /// Outlook for the folder instead of reading a cached listing, and sees the new name immediately.
    /// </para>
    ///
    /// <para>
    /// It is also enormously cheaper: the walk is O(depth) instead of a depth-first search of every
    /// folder in every store, which on a real mailbox is thousands of COM calls per lookup.
    /// </para>
    ///
    /// <para>
    /// Returns null rather than throwing when any segment is missing, so the caller can fall back to
    /// the enumerating search - which still handles a bare folder name, and a path given without its
    /// store prefix.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? WalkFolderPath(Outlook.NameSpace session, string normalizedPath)
    {
        string[] segments = normalizedPath.Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0)
        {
            return null;
        }

        Outlook.MAPIFolder? current = null;
        try
        {
            Outlook.Folders? rootFolders = null;
            try
            {
                rootFolders = session.Folders;
                current = TryGetChild(rootFolders, segments[0]);
            }
            finally
            {
                ReleaseComObject(ref rootFolders);
            }

            if (current == null)
            {
                return null;
            }

            for (int index = 1; index < segments.Length; index++)
            {
                Outlook.Folders? children = null;
                Outlook.MAPIFolder? next = null;
                try
                {
                    children = current.Folders;
                    next = TryGetChild(children, segments[index]);
                }
                finally
                {
                    ReleaseComObject(ref children);
                }

                if (next == null)
                {
                    return null;
                }

                ReleaseComObject(ref current);
                current = next;
            }

            var resolved = current;
            current = null;
            return resolved;
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            ReleaseComObject(ref current);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? TryGetChild(Outlook.Folders folders, string name)
    {
        try
        {
            return folders[name];
        }
        catch (COMException)
        {
            // Outlook raises rather than returning null when there is no such child.
            return null;
        }
        catch (ArgumentException)
        {
            return null;
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

    /// <summary>
    /// Normalizes a body's line endings and trims it, without shortening it.
    /// <para>
    /// Kept separate from <see cref="NormalizeBodyPreview"/> because the two have opposite purposes.
    /// A preview is deliberately short: it is for a human or an LLM to read. Searching is the
    /// opposite - shortening the text before matching it silently loses hits and reports them as
    /// "no such mail", so search must see the whole body.
    /// </para>
    /// </summary>
    internal static string NormalizeBodyText(string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return string.Empty;
        }

        return body
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim();
    }

    internal static string NormalizeBodyPreview(string? body, int maxLength = 500)
    {
        string normalized = NormalizeBodyText(body);

        if (normalized.Length <= maxLength)
        {
            return normalized;
        }

        return normalized[..maxLength].TrimEnd() + "...";
    }

    /// <summary>
    /// Returns a folder's path, or <see langword="null"/> when it does not have a usable one.
    ///
    /// <para>
    /// The null case is not merely defensive. <c>Store.GetDefaultFolder</c> hands back a folder
    /// object for a role the store does not actually have - an online archive with no Inbox still
    /// answers for <c>olFolderInbox</c> - and that object is not parented in the store's tree. Its
    /// <c>FolderPath</c> then degenerates to the folder's entry id: a long hex string that looks
    /// like a value but cannot be passed back as a <c>folder</c> argument, because folder resolution
    /// searches the tree by name. Verified against a real Exchange archive, where nine of ten
    /// default roles behaved this way and only the one folder that genuinely exists returned a path.
    /// </para>
    ///
    /// <para>
    /// Returning that hex string would be the project's characteristic failure: a confident-looking
    /// answer the caller cannot use and has no way to recognise as wrong. A real Outlook folder path
    /// always begins with <c>\\</c>, so anything that does not is reported as absent instead. See #38.
    /// </para>
    /// </summary>
    internal static string? GetFolderPath(Outlook.MAPIFolder? folder)
    {
        if (folder == null)
        {
            return null;
        }

        try
        {
            string? path = folder.FolderPath;
            return path != null && path.StartsWith(@"\\", StringComparison.Ordinal) ? path : null;
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
