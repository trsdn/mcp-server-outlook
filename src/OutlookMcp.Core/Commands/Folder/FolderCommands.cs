using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Folder;

public class FolderCommands : IFolderCommands
{
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> DefaultFolderAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["inbox"] = Outlook.OlDefaultFolders.olFolderInbox,
            ["drafts"] = Outlook.OlDefaultFolders.olFolderDrafts,
            ["sent"] = Outlook.OlDefaultFolders.olFolderSentMail,
            ["outbox"] = Outlook.OlDefaultFolders.olFolderOutbox,
            ["deleted"] = Outlook.OlDefaultFolders.olFolderDeletedItems,
            ["calendar"] = Outlook.OlDefaultFolders.olFolderCalendar,
            ["contacts"] = Outlook.OlDefaultFolders.olFolderContacts,
            ["tasks"] = Outlook.OlDefaultFolders.olFolderTasks,
            ["notes"] = Outlook.OlDefaultFolders.olFolderNotes,
            ["junk"] = Outlook.OlDefaultFolders.olFolderJunk
        };

    /// <summary>
    /// The folder roles Outlook will open in another person's mailbox. This is deliberately narrower
    /// than <see cref="DefaultFolderAliases"/>: <c>GetSharedDefaultFolder</c> accepts only these, and
    /// passing anything else produces a COM error a caller cannot act on. Rejecting up front lets the
    /// error name the supported set instead.
    /// </summary>
    private static readonly Dictionary<string, Outlook.OlDefaultFolders> SharedFolderRoles =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["inbox"] = Outlook.OlDefaultFolders.olFolderInbox,
            ["calendar"] = Outlook.OlDefaultFolders.olFolderCalendar,
            ["contacts"] = Outlook.OlDefaultFolders.olFolderContacts,
            ["tasks"] = Outlook.OlDefaultFolders.olFolderTasks,
            ["notes"] = Outlook.OlDefaultFolders.olFolderNotes,
            ["journal"] = Outlook.OlDefaultFolders.olFolderJournal
        };

    private static readonly (string Role, Outlook.OlDefaultFolders Folder)[] DefaultFolders =
    [
        ("inbox", Outlook.OlDefaultFolders.olFolderInbox),
        ("drafts", Outlook.OlDefaultFolders.olFolderDrafts),
        ("sent", Outlook.OlDefaultFolders.olFolderSentMail),
        ("outbox", Outlook.OlDefaultFolders.olFolderOutbox),
        ("deleted", Outlook.OlDefaultFolders.olFolderDeletedItems),
        ("calendar", Outlook.OlDefaultFolders.olFolderCalendar),
        ("contacts", Outlook.OlDefaultFolders.olFolderContacts),
        ("tasks", Outlook.OlDefaultFolders.olFolderTasks),
        ("notes", Outlook.OlDefaultFolders.olFolderNotes),
        ("junk", Outlook.OlDefaultFolders.olFolderJunk)
    ];

    /// <summary>
    /// Enumerates the default folder roles, optionally from a specific store (#38).
    ///
    /// <para>
    /// <c>NameSpace.GetDefaultFolder</c> always reads the default delivery store, so without
    /// <paramref name="storeId"/> this reports one mailbox's folders and says nothing about the
    /// others existing. <c>Store.GetDefaultFolder</c> is the per-store equivalent. Note that the two
    /// take different enums - <c>OlDefaultFolders</c> against the session, <c>OlDefaultFolders</c>
    /// against the store as well, but a store may legitimately refuse a role it does not have (a PST
    /// has no Outbox, an archive has no Junk), which is reported as <c>available: false</c> rather
    /// than as a failure.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderListResult ListDefault(bool includeItemCounts = false, string? storeId = null)
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderListDefault",
            (application, session) =>
            {
                Outlook.Store? targetStore = null;

                try
                {
                    if (!string.IsNullOrWhiteSpace(storeId))
                    {
                        targetStore = FindStore(session, storeId!);

                        // A store id that does not resolve must fail. Falling back to the default
                        // store would hand back real folders with real item counts from a mailbox
                        // the caller did not ask for, under success: true.
                        if (targetStore == null)
                        {
                            return new OutlookFolderListResult
                            {
                                Success = false,
                                ErrorMessage =
                                    $"No store in this Outlook profile has the id '{storeId}'. "
                                    + "Use folder list-stores to discover the available store ids."
                            };
                        }
                    }

                    var result = new OutlookFolderListResult
                    {
                        Success = true
                    };

                    string? targetStoreId = targetStore != null ? SafeGet(() => targetStore.StoreID) : null;
                    string? targetStoreName = targetStore != null ? SafeGet(() => targetStore.DisplayName) : null;

                    foreach (var entry in DefaultFolders)
                    {
                        Outlook.MAPIFolder? folder = null;
                        Outlook.Items? items = null;
                        Outlook.Store? owningStore = null;

                        try
                        {
                            folder = targetStore != null
                                ? targetStore.GetDefaultFolder(entry.Folder)
                                : session.GetDefaultFolder(entry.Folder);

                            int? itemCount = null;
                            if (includeItemCounts)
                            {
                                items = folder.Items;
                                itemCount = items.Count;
                            }

                            string? folderStoreId = targetStoreId;
                            string? folderStoreName = targetStoreName;

                            if (folderStoreId == null)
                            {
                                owningStore = SafeGetStore(folder);
                                folderStoreId = owningStore != null ? SafeGet(() => owningStore.StoreID) : null;
                                folderStoreName = owningStore != null ? SafeGet(() => owningStore.DisplayName) : null;
                            }

                            string? folderPath = OutlookInteropRunner.GetFolderPath(folder);

                            // A store answers for every default role whether or not it has one. An
                            // online archive with no Inbox still returns a folder object for
                            // olFolderInbox - but that object is not in the store's tree and has no
                            // path, so nothing in this surface can address it. Reporting it as
                            // available would be a confident answer to a question the store never
                            // actually answered. See #38.
                            if (folderPath == null)
                            {
                                result.Folders.Add(new OutlookFolderInfo
                                {
                                    Role = entry.Role,
                                    Available = false,
                                    StoreId = folderStoreId,
                                    StoreName = folderStoreName,
                                    Note =
                                        "This store does not have a folder in this role. Outlook returns a "
                                        + "placeholder that is not in the store's folder tree and cannot be "
                                        + "addressed. Use folder list-children on the store's root folder to see "
                                        + "what it does contain."
                                });
                                continue;
                            }

                            result.Folders.Add(new OutlookFolderInfo
                            {
                                Role = entry.Role,
                                Available = true,
                                Name = folder.Name,
                                FolderPath = folderPath,
                                StoreId = folderStoreId,
                                StoreName = folderStoreName,
                                ItemCount = itemCount
                            });
                        }
                        catch
                        {
                            result.Folders.Add(new OutlookFolderInfo
                            {
                                Role = entry.Role,
                                Available = false,
                                StoreId = targetStoreId,
                                StoreName = targetStoreName
                            });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref owningStore);
                            OutlookInteropRunner.ReleaseComObject(ref items);
                            OutlookInteropRunner.ReleaseComObject(ref folder);
                        }
                    }

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref targetStore);
                }
            },
            ex => new OutlookFolderListResult
            {
                Success = false,
                ErrorMessage = $"Failed to read Outlook default folders: {ex.Message}"
            });
    }

    /// <summary>
    /// Enumerates every store in the profile, so a caller can discover mailboxes other than the
    /// default delivery store (#38).
    ///
    /// <para>
    /// Accounts are read alongside the stores rather than exposed separately. An account is only
    /// interesting here for the address it delivers to, and matching it onto its store here means a
    /// caller never has to correlate two lists by id to answer "which mailbox is this?". Stores with
    /// no delivering account - archives, imported data files - are still listed, with the account
    /// fields absent rather than invented.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookStoreListResult ListStores()
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderListStores",
            (application, session) =>
            {
                Outlook.Stores? stores = null;
                Outlook.Store? defaultStore = null;

                try
                {
                    var result = new OutlookStoreListResult { Success = true };

                    Dictionary<string, (string? Smtp, string? Name)> accountsByStoreId = ReadAccounts(session);

                    defaultStore = SafeGetDefaultStore(session);
                    string? defaultStoreId = defaultStore != null ? SafeGet(() => defaultStore.StoreID) : null;

                    stores = session.Stores;
                    int count = stores.Count;

                    for (int index = 1; index <= count; index++)
                    {
                        Outlook.Store? store = null;
                        Outlook.MAPIFolder? root = null;

                        try
                        {
                            store = stores[index];
                            string? id = SafeGet(() => store.StoreID);

                            // A store with no id cannot be addressed, so listing it would advertise
                            // something unreachable.
                            if (string.IsNullOrWhiteSpace(id))
                            {
                                continue;
                            }

                            root = SafeGetRootFolder(store);

                            accountsByStoreId.TryGetValue(id!, out var account);

                            result.Stores.Add(new OutlookStoreInfo
                            {
                                StoreId = id!,
                                DisplayName = SafeGet(() => store.DisplayName) ?? "(unnamed store)",
                                IsDefaultStore = defaultStoreId != null
                                    && string.Equals(id, defaultStoreId, StringComparison.Ordinal),
                                IsDataFileStore = SafeGetBool(() => store.IsDataFileStore),
                                ExchangeStoreType = DescribeExchangeStoreType(store),
                                FilePath = SafeGet(() => store.FilePath),
                                AccountSmtpAddress = account.Smtp,
                                AccountDisplayName = account.Name,
                                RootFolderPath = root != null ? OutlookInteropRunner.GetFolderPath(root) : null
                            });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref root);
                            OutlookInteropRunner.ReleaseComObject(ref store);
                        }
                    }

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref defaultStore);
                    OutlookInteropRunner.ReleaseComObject(ref stores);
                }
            },
            ex => new OutlookStoreListResult
            {
                Success = false,
                ErrorMessage = $"Failed to read Outlook stores: {ex.Message}"
            });
    }

    /// <summary>
    /// Opens another person's default folder by address, using delegate or shared-mailbox rights
    /// (#38).
    ///
    /// <para>
    /// <c>NameSpace.GetSharedDefaultFolder</c> reaches a mailbox that is not in the profile at all,
    /// which is the only way to read a colleague's calendar without adding their account. It takes a
    /// <c>Recipient</c>, so the address has to be resolved against the address book first.
    /// </para>
    ///
    /// <para>
    /// <b>The unresolved-recipient case is the dangerous one.</b> <c>Resolve</c> returns
    /// <see langword="false"/> rather than throwing, and Outlook then treats the unresolved recipient
    /// as the current user - so a caller asking for a colleague's calendar would be shown their own,
    /// with <c>success: true</c> and a perfectly plausible folder. An unresolved address is therefore
    /// refused outright.
    /// </para>
    ///
    /// <para>
    /// <b><c>Resolve</c> is not an existence check, though.</b> Verified live: it returns
    /// <see langword="true"/> for <c>no-such-person@invalid.example</c>, because Outlook accepts any
    /// syntactically valid SMTP address as a one-off recipient without consulting the directory. So
    /// the guard above catches only part of the problem; the rest is caught by
    /// <c>GetSharedDefaultFolder</c> itself failing, which is why that failure is reported with the
    /// mailbox and role named rather than as a bare HRESULT. Do not tighten this into a directory
    /// lookup - on a non-Exchange profile every legitimate recipient is a plain SMTP entry too, so
    /// that would reject real mailboxes.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult OpenShared(string? address = null, string? role = null)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            return new OutlookFolderResolveResult
            {
                Success = false,
                ErrorMessage =
                    "An address is required: open-shared answers the question of whose mailbox to "
                    + "open, so there is no sensible default."
            };
        }

        string requestedRole = string.IsNullOrWhiteSpace(role) ? "inbox" : role!.Trim();

        if (!SharedFolderRoles.TryGetValue(requestedRole, out var folderType))
        {
            return new OutlookFolderResolveResult
            {
                Success = false,
                RequestedFolder = requestedRole,
                ErrorMessage =
                    $"'{requestedRole}' is not a folder role that can be opened in another mailbox. "
                    + "Outlook supports: " + string.Join(", ", SharedFolderRoles.Keys.Order(StringComparer.Ordinal))
                    + "."
            };
        }

        return OutlookInteropRunner.Execute(
            "OutlookFolderOpenShared",
            (application, session) =>
            {
                Outlook.Recipient? recipient = null;
                Outlook.MAPIFolder? folder = null;
                Outlook.Folders? children = null;
                Outlook.Items? items = null;
                Outlook.Store? store = null;

                try
                {
                    recipient = session.CreateRecipient(address);

                    if (!recipient.Resolve())
                    {
                        return new OutlookFolderResolveResult
                        {
                            Success = false,
                            RequestedFolder = requestedRole,
                            ErrorMessage =
                                $"Outlook could not resolve '{address}' against the address book. The "
                                + "request is refused rather than served, because an unresolved recipient "
                                + "makes Outlook return the signed-in user's own folder instead - which "
                                + "would look like a successful answer about someone else's mailbox."
                        };
                    }

                    folder = session.GetSharedDefaultFolder(recipient, folderType);

                    string? folderPath = OutlookInteropRunner.GetFolderPath(folder);

                    if (folderPath == null)
                    {
                        return new OutlookFolderResolveResult
                        {
                            Success = false,
                            RequestedFolder = requestedRole,
                            ErrorMessage =
                                $"Outlook returned a '{requestedRole}' folder for '{address}' that has no "
                                + "usable path, so nothing can be read from it. This normally means the "
                                + "mailbox exists but the folder is not actually shared."
                        };
                    }

                    children = folder.Folders;
                    items = folder.Items;
                    store = SafeGetStore(folder);

                    return new OutlookFolderResolveResult
                    {
                        Success = true,
                        Resolved = true,
                        RequestedFolder = requestedRole,
                        Name = SafeGet(() => folder.Name),
                        FolderPath = folderPath,
                        StoreId = store != null ? SafeGet(() => store.StoreID) : null,
                        DefaultRole = requestedRole,
                        ChildFolderCount = children.Count,
                        ItemCount = items.Count
                    };
                }
                catch (COMException ex) when (OutlookInteropRunner.IsObjectModelGuardDenial(ex))
                {
                    return new OutlookFolderResolveResult
                    {
                        Success = false,
                        RequestedFolder = requestedRole,
                        ErrorMessage =
                            "Outlook's security prompt blocked reading the address book, which is needed "
                            + "to resolve the recipient. Answer the prompt and retry."
                    };
                }
                catch (COMException ex)
                {
                    // Outlook reports both "no such mailbox" and "you do not have permission" as a
                    // plain COM failure, and cannot distinguish them: Recipient.Resolve accepts any
                    // syntactically valid SMTP address as a one-off without consulting the directory,
                    // so a typo reaches this point looking exactly like a permissions problem. Naming
                    // both possibilities is what makes the error actionable.
                    return new OutlookFolderResolveResult
                    {
                        Success = false,
                        RequestedFolder = requestedRole,
                        ErrorMessage =
                            $"Could not open the '{requestedRole}' folder of '{address}'. Either no such "
                            + "mailbox exists, or it has not granted access. Outlook cannot tell those "
                            + $"apart. Outlook reported: {ex.Message}"
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref store);
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref children);
                    OutlookInteropRunner.ReleaseComObject(ref folder);
                    OutlookInteropRunner.ReleaseComObject(ref recipient);
                }
            },
            ex => new OutlookFolderResolveResult
            {
                Success = false,
                RequestedFolder = requestedRole,
                ErrorMessage = $"Failed to open a shared Outlook folder: {ex.Message}"
            });
    }

    /// <summary>
    /// Creates a child folder under an existing folder.
    ///
    /// <para>
    /// The name is checked against the existing children first. Outlook does raise its own error for
    /// a duplicate, but only after the fact and phrased in terms of MAPI; checking here means the
    /// caller is told which name collided, in a folder they can then list.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult Create(string? parentFolder = null, string? name = null)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return Refuse(
                "A folder name is required, and it cannot be blank. Outlook accepts a blank name in "
                + "some builds and produces a folder that cannot afterwards be addressed by path.");
        }

        string folderName = name!.Trim();

        // A backslash is the path separator, so a name containing one would produce a folder whose
        // own path cannot be parsed back into it - findable in the UI, unreachable through this tool.
        if (folderName.Contains('\\'))
        {
            return Refuse(
                $"'{folderName}' cannot be used as a folder name: a backslash is the folder path "
                + "separator, so the resulting folder could not be addressed by path afterwards.");
        }

        return OutlookInteropRunner.Execute(
            "OutlookFolderCreate",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? parent = null;
                Outlook.Folders? children = null;
                Outlook.MAPIFolder? created = null;

                try
                {
                    parent = OutlookInteropRunner.ResolveFolder(
                        application, session, parentFolder, DefaultFolderAliases, ref explorer);

                    if (parent == null)
                    {
                        return Refuse($"The parent folder '{parentFolder}' could not be resolved.");
                    }

                    children = parent.Folders;

                    if (ContainsChildNamed(children, folderName))
                    {
                        return Refuse(
                            $"'{parent.Name}' already has a child folder named '{folderName}'. Nothing "
                            + "was created, and the existing folder is deliberately not returned: a "
                            + "caller expecting a new empty folder would otherwise be handed one with "
                            + "contents in it.");
                    }

                    created = children.Add(folderName);

                    return Describe(created, "created");
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref created);
                    OutlookInteropRunner.ReleaseComObject(ref children);
                    OutlookInteropRunner.ReleaseComObject(ref parent);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => Refuse($"Failed to create the Outlook folder: {ex.Message}"));
    }

    /// <summary>
    /// Renames a folder.
    ///
    /// <para>
    /// Refused for a default folder and a store root - see <see cref="RefuseIfProtected"/> for why
    /// that guard exists at all.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult Rename(string? folder = null, string? name = null)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return Refuse("A new folder name is required, and it cannot be blank.");
        }

        string folderName = name!.Trim();

        if (folderName.Contains('\\'))
        {
            return Refuse(
                $"'{folderName}' cannot be used as a folder name: a backslash is the folder path "
                + "separator, so the folder could not be addressed by path afterwards.");
        }

        return OutlookInteropRunner.Execute(
            "OutlookFolderRename",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? target = null;

                try
                {
                    target = OutlookInteropRunner.ResolveFolder(
                        application, session, folder, DefaultFolderAliases, ref explorer);

                    if (target == null)
                    {
                        return Refuse($"The folder '{folder}' could not be resolved.");
                    }

                    var refusal = RefuseIfProtected(session, target, "renamed");
                    if (refusal != null)
                    {
                        return refusal;
                    }

                    string? entryId = SafeGet(() => target.EntryID);

                    target.Name = folderName;

                    // Setting Name does not refresh the reference: reading Name or FolderPath back
                    // from it still gives the old values, so returning them would report a rename
                    // that looks not to have happened. Verified live. Outlook keeps the entry id
                    // across a rename, so re-fetching by id gives the folder in its new state.
                    if (entryId != null)
                    {
                        Outlook.MAPIFolder? refreshed = null;
                        try
                        {
                            refreshed = session.GetFolderFromID(entryId);
                            if (refreshed != null)
                            {
                                var renamed = Describe(refreshed, "renamed");
                                ConfirmPathResolves(session, renamed, "renamed");
                                return renamed;
                            }
                        }
                        catch (COMException)
                        {
                            // Fall through to the stale reference below rather than failing an
                            // operation that has already taken effect.
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref refreshed);
                        }
                    }

                    var stale = Describe(target, "renamed");
                    stale.Note =
                        "The folder was renamed, but Outlook would not return it fresh afterwards, so "
                        + "the name and path reported here may still be the old ones. Re-read the "
                        + "parent with list-children before relying on them.";
                    return stale;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref target);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => Refuse($"Failed to rename the Outlook folder: {ex.Message}"));
    }

    /// <summary>
    /// Moves a folder under a different parent, taking its contents and subfolders with it.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult Move(string? folder = null, string? destinationFolder = null)
    {
        if (string.IsNullOrWhiteSpace(destinationFolder))
        {
            return Refuse("A destination folder is required: move answers the question of where to.");
        }

        return OutlookInteropRunner.Execute(
            "OutlookFolderMove",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? target = null;
                Outlook.MAPIFolder? destination = null;
                Outlook.MAPIFolder? moved = null;

                try
                {
                    target = OutlookInteropRunner.ResolveFolder(
                        application, session, folder, DefaultFolderAliases, ref explorer);

                    if (target == null)
                    {
                        return Refuse($"The folder '{folder}' could not be resolved.");
                    }

                    var refusal = RefuseIfProtected(session, target, "moved");
                    if (refusal != null)
                    {
                        return refusal;
                    }

                    destination = OutlookInteropRunner.ResolveFolder(
                        application, session, destinationFolder, DefaultFolderAliases, ref explorer);

                    if (destination == null)
                    {
                        return Refuse($"The destination folder '{destinationFolder}' could not be resolved.");
                    }

                    // Moving a folder into itself or into its own subtree destroys it in some Outlook
                    // builds and hangs in others. Neither is recoverable, so it is refused here.
                    string? targetPath = OutlookInteropRunner.GetFolderPath(target);
                    string? destinationPath = OutlookInteropRunner.GetFolderPath(destination);

                    if (targetPath != null && destinationPath != null
                        && (destinationPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase)
                            || destinationPath.StartsWith(targetPath + "\\", StringComparison.OrdinalIgnoreCase)))
                    {
                        return Refuse(
                            $"'{targetPath}' cannot be moved into itself or into one of its own "
                            + "subfolders.");
                    }

                    string? movedName = SafeGet(() => target.Name);

                    target.MoveTo(destination);

                    // MoveTo returns void in this interop assembly and leaves the original reference
                    // stale, so the answer is built by finding the folder again under its new parent
                    // rather than by trusting either the old reference or a constructed path.
                    Outlook.Folders? destinationChildren = null;
                    try
                    {
                        destinationChildren = destination.Folders;
                        moved = movedName != null ? FindChildNamed(destinationChildren, movedName) : null;
                    }
                    finally
                    {
                        OutlookInteropRunner.ReleaseComObject(ref destinationChildren);
                    }

                    if (moved == null)
                    {
                        return Refuse(
                            $"Outlook reported no error moving '{movedName}', but it is not under "
                            + $"'{destinationPath}' afterwards, so the move cannot be confirmed.");
                    }

                    var result = Describe(moved, "moved");
                    ConfirmPathResolves(session, result, "moved");
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref moved);
                    OutlookInteropRunner.ReleaseComObject(ref destination);
                    OutlookInteropRunner.ReleaseComObject(ref target);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => Refuse($"Failed to move the Outlook folder: {ex.Message}"));
    }

    /// <summary>
    /// Deletes a folder and everything in it.
    ///
    /// <para>
    /// <b>This is not a recycle-bin operation for every store.</b> In a mail store Outlook moves the
    /// folder to Deleted Items; elsewhere it can be gone outright. Either way the contents go with
    /// it, which is why the confirmation gate and the protected-folder guard below are the substance
    /// of this operation.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult Delete(string? folder = null, bool confirm = false)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            return Refuse(
                "A folder is required. Delete has no default target: falling back to the current "
                + "folder here would delete whatever the user happens to have selected.");
        }

        // Gated before Outlook is reached at all: a refusal must not cost a COM round trip, and
        // nothing about the target can make an unconfirmed folder delete acceptable. See #9 and
        // ConfirmationGate for why folder delete is gated where an item delete is not.
        if (!confirm)
        {
            return Refuse(ConfirmationGate.FolderDelete(folder!));
        }

        return OutlookInteropRunner.Execute(
            "OutlookFolderDelete",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? target = null;

                try
                {
                    target = OutlookInteropRunner.ResolveFolder(
                        application, session, folder, DefaultFolderAliases, ref explorer);

                    if (target == null)
                    {
                        return Refuse($"The folder '{folder}' could not be resolved.");
                    }

                    var refusal = RefuseIfProtected(session, target, "deleted");
                    if (refusal != null)
                    {
                        return refusal;
                    }

                    // Read the description before deleting: afterwards the reference is dead and the
                    // caller would get a success with no indication of what went.
                    var describedBefore = Describe(target, "deleted");

                    target.Delete();

                    return describedBefore;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref target);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => Refuse($"Failed to delete the Outlook folder: {ex.Message}"));
    }

    /// <summary>
    /// The guard that makes folder mutation safe to expose at all.
    ///
    /// <para>
    /// Outlook will delete the Inbox. <c>MAPIFolder.Delete</c> raises no error, shows no prompt, and
    /// takes every message with it; renaming or moving one is less final but breaks every stored path
    /// pointing at it, including Outlook's own. Nothing in COM refuses any of this, so it is refused
    /// here - by comparing entry ids, not names, because a folder called "Inbox" that is not the
    /// default Inbox is an ordinary folder and must stay deletable.
    /// </para>
    ///
    /// <para>
    /// Store roots are refused for a plainer reason: there is no parent to remove one from, and
    /// "delete this folder" against a root means losing a whole mailbox.
    /// </para>
    ///
    /// <para>
    /// The check spans <b>every</b> store, not just the default one - an archive's Deleted Items is
    /// as much a default folder as the primary mailbox's, and #38 made those reachable.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static OutlookFolderResolveResult? RefuseIfProtected(
        Outlook.NameSpace session,
        Outlook.MAPIFolder target,
        string verb)
    {
        string? entryId = SafeGet(() => target.EntryID);
        if (entryId == null)
        {
            return null;
        }

        Outlook.Stores? stores = null;
        try
        {
            stores = session.Stores;

            // Count is snapshotted rather than re-read each iteration: the collection can shift while
            // Outlook syncs, and an index loop that re-reads it ends up addressing a slot that has
            // gone. That is the "Array index out of bounds" failure seen from the Folders collection.
            int storeCount = stores.Count;

            for (int index = 1; index <= storeCount; index++)
            {
                Outlook.Store? store = null;
                Outlook.MAPIFolder? root = null;
                try
                {
                    store = stores[index];
                    root = SafeGetRootFolder(store);

                    if (root != null && string.Equals(SafeGet(() => root.EntryID), entryId, StringComparison.Ordinal))
                    {
                        return Refuse(
                            $"'{SafeGet(() => target.Name)}' is the root of the '{SafeGet(() => store.DisplayName)}' "
                            + $"store and cannot be {verb}. That would mean losing the whole mailbox.");
                    }

                    foreach (var entry in DefaultFolders)
                    {
                        Outlook.MAPIFolder? candidate = null;
                        try
                        {
                            candidate = store.GetDefaultFolder(entry.Folder);
                        }
                        catch (COMException)
                        {
                            // A store need not have every role - an archive typically has almost none.
                            continue;
                        }

                        try
                        {
                            if (string.Equals(SafeGet(() => candidate!.EntryID), entryId, StringComparison.Ordinal))
                            {
                                return Refuse(
                                    $"'{SafeGet(() => target.Name)}' is the '{entry.Role}' folder of the "
                                    + $"'{SafeGet(() => store.DisplayName)}' store and cannot be {verb}. "
                                    + "Outlook would allow it, and everything filed in it would go too. "
                                    + "If the contents are the problem, move them elsewhere and act on "
                                    + "an ordinary folder instead.");
                            }
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref candidate);
                        }
                    }
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref root);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            }
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref stores);
        }

        return null;
    }

    /// <summary>
    /// Checks that a path this operation just produced actually resolves - and says so in the result
    /// if it does not.
    ///
    /// <para>
    /// This exists because rename and move both used to hand back a path that did not work. The cause
    /// was not a delay in Outlook: <c>NameSpace.Folders</c> enumeration goes stale after a rename and
    /// keeps reporting the old name for the life of the process, so a lookup that walked it could
    /// never find the new path however long it waited. An earlier version of this method retried on
    /// the assumption that Outlook was lagging; it failed just as often, which is what exposed the
    /// real cause. The fix is in <c>OutlookInteropRunner.WalkFolderPath</c>, which asks Outlook for
    /// each segment by name instead of reading a cached listing.
    /// </para>
    ///
    /// <para>
    /// The check is kept, without the retry, as a post-condition. Returning a path that reports
    /// success but does not work is this project's characteristic failure, and one lookup is a cheap
    /// guard against it coming back. Retrying is deliberately not reinstated: it would only hide the
    /// next instance the way it hid this one.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static void ConfirmPathResolves(
        Outlook.NameSpace session,
        OutlookFolderResolveResult result,
        string verb)
    {
        if (result.FolderPath == null)
        {
            return;
        }

        Outlook.MAPIFolder? found = null;
        try
        {
            Outlook.Explorer? none = null;
            found = OutlookInteropRunner.ResolveFolder(
                session.Application, session, result.FolderPath, DefaultFolderAliases, ref none);

            if (found != null)
            {
                return;
            }
        }
        catch (COMException)
        {
            // Treated the same as not found: the note below tells the caller not to trust the path.
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref found);
        }

        result.Note =
            $"The folder was {verb}, but Outlook does not resolve "
            + $"'{result.FolderPath}' by path. Re-read the parent with list-children before using it.";
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static bool ContainsChildNamed(Outlook.Folders children, string name)
    {
        Outlook.MAPIFolder? found = FindChildNamed(children, name);
        try
        {
            return found != null;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref found);
        }
    }

    /// <summary>
    /// Finds a child by name.
    ///
    /// <para>
    /// Uses Outlook's own name indexer rather than walking the collection by position. Walking it was
    /// tried first and fails intermittently with "Array index out of bounds" - the <c>Folders</c>
    /// collection can shift under an index loop while Outlook syncs, and the loop is then reading a
    /// slot that no longer exists. The name indexer has no such window. It is case-insensitive, which
    /// matches Outlook's own rule that two siblings cannot differ only in case.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? FindChildNamed(Outlook.Folders children, string name)
    {
        try
        {
            return children[name];
        }
        catch (COMException)
        {
            // Outlook raises rather than returning null when the name is not present.
            return null;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    /// <summary>
    /// Builds the answer for a folder that was just created, renamed, moved or deleted, from the
    /// folder itself rather than from the arguments the caller passed.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static OutlookFolderResolveResult Describe(Outlook.MAPIFolder folder, string verb)
    {
        Outlook.Folders? children = null;
        Outlook.Items? items = null;
        Outlook.Store? store = null;

        try
        {
            children = folder.Folders;
            items = folder.Items;
            store = SafeGetStore(folder);

            string? path = OutlookInteropRunner.GetFolderPath(folder);

            return new OutlookFolderResolveResult
            {
                Success = true,
                Resolved = true,
                Name = SafeGet(() => folder.Name),
                FolderPath = path,
                StoreId = store != null ? SafeGet(() => store.StoreID) : null,
                ChildFolderCount = children.Count,
                ItemCount = items.Count,
                Note = path == null
                    ? $"The folder was {verb}, but Outlook reports no usable path for it, so it cannot "
                        + "be addressed by path afterwards."
                    : null
            };
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref store);
            OutlookInteropRunner.ReleaseComObject(ref items);
            OutlookInteropRunner.ReleaseComObject(ref children);
        }
    }

    private static OutlookFolderResolveResult Refuse(string message) => new()
    {
        Success = false,
        Resolved = false,
        ErrorMessage = message
    };

    /// <summary>
    /// Maps store id to the account that delivers there. Comparison is ordinal: a store id is an
    /// opaque hex string, so case-insensitive matching would risk conflating two distinct stores.
    /// </summary>    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Dictionary<string, (string? Smtp, string? Name)> ReadAccounts(Outlook.NameSpace session)
    {
        var map = new Dictionary<string, (string? Smtp, string? Name)>(StringComparer.Ordinal);
        Outlook.Accounts? accounts = null;

        try
        {
            accounts = session.Accounts;
            int count = accounts.Count;

            for (int index = 1; index <= count; index++)
            {
                Outlook.Account? account = null;
                Outlook.Store? deliveryStore = null;

                try
                {
                    account = accounts[index];
                    deliveryStore = SafeGetDeliveryStore(account);
                    string? id = deliveryStore != null ? SafeGet(() => deliveryStore.StoreID) : null;

                    if (!string.IsNullOrWhiteSpace(id))
                    {
                        map[id!] = (SafeGet(() => account.SmtpAddress), SafeGet(() => account.DisplayName));
                    }
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref deliveryStore);
                    OutlookInteropRunner.ReleaseComObject(ref account);
                }
            }
        }
        catch (COMException)
        {
            // A profile that will not enumerate accounts still has stores worth listing. The account
            // fields are simply absent, which is honest; inventing them would not be.
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref accounts);
        }

        return map;
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Store? FindStore(Outlook.NameSpace session, string storeId)
    {
        Outlook.Stores? stores = null;

        try
        {
            stores = session.Stores;
            int count = stores.Count;

            for (int index = 1; index <= count; index++)
            {
                Outlook.Store? store = null;
                bool keep = false;

                try
                {
                    store = stores[index];
                    keep = string.Equals(SafeGet(() => store.StoreID), storeId, StringComparison.Ordinal);

                    if (keep)
                    {
                        return store;
                    }
                }
                finally
                {
                    if (!keep)
                    {
                        OutlookInteropRunner.ReleaseComObject(ref store);
                    }
                }
            }

            return null;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref stores);
        }
    }

    /// <summary>
    /// Renders <c>OlExchangeStoreType</c> as a stable camelCase name. The raw integer would be a
    /// number a caller has to look up, and the enum name itself carries a Hungarian prefix.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? DescribeExchangeStoreType(Outlook.Store store)
    {
        try
        {
            return store.ExchangeStoreType switch
            {
                Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox => "primaryExchangeMailbox",
                Outlook.OlExchangeStoreType.olExchangeMailbox => "exchangeMailbox",
                Outlook.OlExchangeStoreType.olExchangePublicFolder => "exchangePublicFolder",
                Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox => "additionalExchangeMailbox",
                Outlook.OlExchangeStoreType.olNotExchange => "notExchange",
                _ => null
            };
        }
        catch (COMException)
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Store? SafeGetDefaultStore(Outlook.NameSpace session)
    {
        Outlook.MAPIFolder? inbox = null;

        try
        {
            // There is no NameSpace.DefaultStore. The default delivery store is by definition the one
            // holding the default Inbox, which is exactly what GetDefaultFolder returns.
            inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            return SafeGetStore(inbox);
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            OutlookInteropRunner.ReleaseComObject(ref inbox);
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Store? SafeGetStore(Outlook.MAPIFolder folder)
    {
        try
        {
            return folder.Store;
        }
        catch (COMException)
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.MAPIFolder? SafeGetRootFolder(Outlook.Store store)
    {
        try
        {
            return store.GetRootFolder();
        }
        catch (COMException)
        {
            return null;
        }
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static Outlook.Store? SafeGetDeliveryStore(Outlook.Account account)
    {
        try
        {
            return account.DeliveryStore;
        }
        catch (COMException)
        {
            return null;
        }
    }


    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderListResult ListChildren(
        string? parentFolder = null,
        bool includeItemCounts = false)
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderListChildren",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Folders? childFolders = null;

                try
                {
                    resolvedFolder = OutlookInteropRunner.ResolveFolder(
                        application,
                        session,
                        parentFolder,
                        DefaultFolderAliases,
                        ref explorer);

                    if (resolvedFolder == null)
                    {
                        return new OutlookFolderListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(parentFolder)
                        };
                    }

                    var result = new OutlookFolderListResult
                    {
                        Success = true
                    };

                    childFolders = resolvedFolder.Folders;
                    int childCount = childFolders.Count;
                    for (int index = 1; index <= childCount; index++)
                    {
                        Outlook.MAPIFolder? childFolder = null;
                        Outlook.Items? items = null;

                        try
                        {
                            childFolder = childFolders[index];
                            int? itemCount = null;
                            if (includeItemCounts)
                            {
                                items = childFolder.Items;
                                itemCount = items.Count;
                            }

                            result.Folders.Add(new OutlookFolderInfo
                            {
                                Role = SafeGet(() => childFolder.Name) ?? $"child-{index}",
                                Available = true,
                                Name = SafeGet(() => childFolder.Name),
                                FolderPath = OutlookInteropRunner.GetFolderPath(childFolder),
                                ItemCount = itemCount
                            });
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref items);
                            OutlookInteropRunner.ReleaseComObject(ref childFolder);
                        }
                    }

                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref childFolders);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookFolderListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook child folders: {ex.Message}"
            });
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

    private static string BuildUnknownFolderMessage(string? folder)
    {
        const string supportedFolders = "current, inbox, drafts, sent, outbox, deleted, calendar, contacts, tasks, notes, junk, or an Outlook folder path";
        return string.IsNullOrWhiteSpace(folder)
            ? $"Could not resolve the Outlook folder. Supported folder values: {supportedFolders}."
            : $"Unsupported Outlook folder '{folder}'. Supported folder values: {supportedFolders}.";
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderResolveResult ResolvePath(
        string? folder = null,
        bool includeItemCount = true)
    {
        return OutlookInteropRunner.Execute(
            "OutlookFolderResolvePath",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Folders? childFolders = null;
                Outlook.Items? items = null;

                try
                {
                    resolvedFolder = OutlookInteropRunner.ResolveFolder(
                        application,
                        session,
                        folder,
                        DefaultFolderAliases,
                        ref explorer);

                    if (resolvedFolder == null)
                    {
                        return new OutlookFolderResolveResult
                        {
                            Success = false,
                            RequestedFolder = folder,
                            Resolved = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    childFolders = resolvedFolder.Folders;
                    int? itemCount = null;
                    if (includeItemCount)
                    {
                        items = resolvedFolder.Items;
                        itemCount = SafeGetInt(() => items.Count);
                    }

                    return new OutlookFolderResolveResult
                    {
                        Success = true,
                        RequestedFolder = folder,
                        Resolved = true,
                        Name = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        StoreId = SafeGet(() => resolvedFolder.StoreID),
                        DefaultRole = TryGetDefaultRole(folder),
                        ChildFolderCount = SafeGetInt(() => childFolders.Count),
                        ItemCount = itemCount
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref childFolders);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookFolderResolveResult
            {
                Success = false,
                RequestedFolder = folder,
                Resolved = false,
                ErrorMessage = $"Failed to resolve the Outlook folder: {ex.Message}"
            });
    }

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookFolderItemListResult ListItems(
        string? folder = null,
        int maxCount = 25,
        bool includePreview = false)
    {
        int boundedMaxCount = Math.Clamp(maxCount, 1, 100);

        return OutlookInteropRunner.Execute(
            "OutlookFolderListItems",
            (application, session) =>
            {
                Outlook.Explorer? explorer = null;
                Outlook.MAPIFolder? resolvedFolder = null;
                Outlook.Items? items = null;

                try
                {
                    resolvedFolder = OutlookInteropRunner.ResolveFolder(
                        application,
                        session,
                        folder,
                        DefaultFolderAliases,
                        ref explorer);

                    if (resolvedFolder == null)
                    {
                        return new OutlookFolderItemListResult
                        {
                            Success = false,
                            ErrorMessage = BuildUnknownFolderMessage(folder)
                        };
                    }

                    items = resolvedFolder.Items;
                    int totalItemCount = SafeGetInt(() => items.Count);

                    // Without this the cap below returns an arbitrary subset in store order, which
                    // a caller reads as "this is what is in the folder" (#91). Ordering is attempted
                    // newest-first and the property actually used is reported, because a folder of
                    // appointments or contacts has no received time and the honest answer there is a
                    // different ordering rather than a pretended one.
                    string? sortedBy = TrySortNewestFirst(items);

                    var result = new OutlookFolderItemListResult
                    {
                        Success = true,
                        FolderName = SafeGet(() => resolvedFolder.Name),
                        FolderPath = OutlookInteropRunner.GetFolderPath(resolvedFolder),
                        TotalItemCount = totalItemCount,
                        Truncated = totalItemCount > boundedMaxCount,
                        SortedBy = sortedBy,
                        SortDirection = sortedBy == null ? null : "descending"
                    };

                    for (int index = 1; index <= totalItemCount && result.Items.Count < boundedMaxCount; index++)
                    {
                        object? rawItem = null;

                        try
                        {
                            rawItem = items[index];
                            var info = CreateFolderItemInfo(rawItem, includePreview);
                            if (info != null)
                            {
                                result.Items.Add(info);
                            }
                        }
                        finally
                        {
                            OutlookInteropRunner.ReleaseComObject(ref rawItem);
                        }
                    }

                    result.ReturnedCount = result.Items.Count;
                    return result;
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref items);
                    OutlookInteropRunner.ReleaseComObject(ref resolvedFolder);
                    OutlookInteropRunner.ReleaseComObject(ref explorer);
                }
            },
            ex => new OutlookFolderItemListResult
            {
                Success = false,
                ErrorMessage = $"Failed to enumerate Outlook folder items: {ex.Message}"
            });
    }

    /// <summary>
    /// Orders a folder's items newest-first, returning the property that worked.
    ///
    /// <para>
    /// <c>ReceivedTime</c> is preferred because it is what "newest" means for mail, but it does not
    /// exist on appointments, contacts or tasks and Outlook throws rather than ignoring it there.
    /// <c>LastModificationTime</c> exists on every item type and is the fallback. If both fail the
    /// caller is told the order is unknown rather than being handed store order dressed up as an
    /// ordering. See #91.
    /// </para>
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string? TrySortNewestFirst(Outlook.Items items)
    {
        foreach ((string property, string reported) in SortCandidates)
        {
            try
            {
                items.Sort(property, true);
                return reported;
            }
            catch (COMException)
            {
                // Property not available on this folder's item types; try the next one.
            }
        }

        return null;
    }

    private static readonly (string Property, string Reported)[] SortCandidates =
    [
        ("[ReceivedTime]", "receivedTime"),
        ("[LastModificationTime]", "lastModificationTime")
    ];

    private static DateTimeOffset? SafeGetDateTimeOffset(Func<DateTime> getter)
    {
        try
        {
            DateTime value = getter();
            if (value == default)
            {
                return null;
            }

            return new DateTimeOffset(value);
        }
        catch
        {
            return null;
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

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static OutlookFolderItemInfo? CreateFolderItemInfo(object rawItem, bool includePreview)
    {
        if (rawItem is Outlook.MailItem mail)
        {
            return new OutlookFolderItemInfo
            {
                EntryId = SafeGet(() => mail.EntryID),
                StoreId = SafeGet(() => (mail.Parent as Outlook.MAPIFolder)?.StoreID),
                ItemType = "mail",
                MessageClass = SafeGet(() => mail.MessageClass),
                Subject = SafeGet(() => mail.Subject),
                Name = SafeGet(() => mail.SenderName),
                Preview = includePreview
                    ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => mail.Body))
                    : null,
                ReceivedTime = SafeGetDateTimeOffset(() => mail.ReceivedTime),
                Unread = SafeGetBool(() => mail.UnRead)
            };
        }

        if (rawItem is Outlook.AppointmentItem appointment)
        {
            return new OutlookFolderItemInfo
            {
                EntryId = SafeGet(() => appointment.EntryID),
                StoreId = SafeGet(() => (appointment.Parent as Outlook.MAPIFolder)?.StoreID),
                ItemType = "appointment",
                MessageClass = SafeGet(() => appointment.MessageClass),
                Subject = SafeGet(() => appointment.Subject),
                Name = SafeGet(() => appointment.Organizer),
                Preview = includePreview
                    ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => appointment.Body))
                    : null,
                Start = SafeGetDateTimeOffset(() => appointment.Start),
                End = SafeGetDateTimeOffset(() => appointment.End)
            };
        }

        if (rawItem is Outlook.ContactItem contact)
        {
            return new OutlookFolderItemInfo
            {
                EntryId = SafeGet(() => contact.EntryID),
                StoreId = SafeGet(() => (contact.Parent as Outlook.MAPIFolder)?.StoreID),
                ItemType = "contact",
                MessageClass = SafeGet(() => contact.MessageClass),
                Subject = SafeGet(() => contact.CompanyName),
                Name = SafeGet(() => contact.FullName),
                Preview = includePreview
                    ? OutlookInteropRunner.NormalizeBodyPreview(SafeGet(() => contact.Body))
                    : null
            };
        }

        // Reason: rawItem is an Outlook item of a type this method could not identify - a PostItem,
        // JournalItem, DistListItem, or an item from a third-party add-in. The PIA models these as
        // unrelated COM classes with no common interface exposing MessageClass, Subject, FullName or
        // Name, so late binding is the only way to read them. SafeGet swallows the resulting
        // RuntimeBinderException when a given type does not have the member.
        dynamic untypedItem = rawItem;

        return new OutlookFolderItemInfo
        {
            ItemType = SafeGet(() => rawItem.GetType().Name),
            MessageClass = SafeGet(() => (string?)untypedItem.MessageClass),
            Subject = SafeGet(() => (string?)untypedItem.Subject),
            Name = SafeGet(() => (string?)untypedItem.FullName) ?? SafeGet(() => (string?)untypedItem.Name)
        };
    }

    private static string? TryGetDefaultRole(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            return "current";
        }

        return DefaultFolderAliases.ContainsKey(folder) ? folder.ToLowerInvariant() : null;
    }
}
