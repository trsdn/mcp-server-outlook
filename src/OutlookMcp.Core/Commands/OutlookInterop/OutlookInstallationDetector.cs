using System.Runtime.Versioning;
using System.Security.Principal;
using Microsoft.Win32;

namespace OutlookMcp.Core.Commands.OutlookInterop;

/// <summary>
/// The Outlook desktop flavour detected on this machine. The COM object model
/// (<c>Microsoft.Office.Interop.Outlook</c>, ProgID <c>Outlook.Application</c>) exists only in
/// classic Outlook for Windows. The new Outlook for Windows (packaged as
/// <c>Microsoft.OutlookForWindows</c>) has no COM object model at all, so its presence alone does
/// not make this server usable. See issue #35.
/// </summary>
public enum OutlookFlavor
{
    /// <summary>Neither classic Outlook nor new Outlook could be detected on this machine.</summary>
    NotInstalled,

    /// <summary>Classic Outlook for Windows is registered (the <c>Outlook.Application</c> COM ProgID resolves). This is the only flavour this server supports.</summary>
    ClassicDesktop,

    /// <summary>Only the new Outlook for Windows (no COM object model) was detected; classic Outlook is not registered.</summary>
    NewOutlookOnly,

    /// <summary>Detection could not determine a flavour (e.g. non-Windows or registry access denied).</summary>
    Unknown
}

/// <summary>
/// Detects which Outlook desktop flavour (if any) is installed and whether this process is
/// running at a different integrity level than a running classic Outlook would need, without
/// requiring Outlook to be running. Pure registry/OS inspection with no COM dependency, so it is
/// safe to unit test directly (see Rule 30's exception for algorithmic utilities with zero COM
/// dependency).
/// </summary>
public static class OutlookInstallationDetector
{
    private const string NewOutlookPackageFamilyName = "Microsoft.OutlookForWindows";

    /// <summary>
    /// Detects the Outlook flavour installed on this machine by inspecting the
    /// <c>Outlook.Application</c> COM ProgID registration and, if absent, whether the new Outlook
    /// for Windows package is present instead.
    /// </summary>
    public static OutlookFlavor DetectFlavor()
    {
        if (!OperatingSystem.IsWindows())
        {
            return OutlookFlavor.Unknown;
        }

        try
        {
            if (Type.GetTypeFromProgID("Outlook.Application") != null)
            {
                return OutlookFlavor.ClassicDesktop;
            }

            return IsNewOutlookPackagePresent() ? OutlookFlavor.NewOutlookOnly : OutlookFlavor.NotInstalled;
        }
        catch
        {
            return OutlookFlavor.Unknown;
        }
    }

    /// <summary>
    /// Returns true if the new Outlook for Windows package (<c>Microsoft.OutlookForWindows</c>,
    /// no COM object model) appears to be registered for the current user.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static bool IsNewOutlookPackagePresent()
    {
        try
        {
            using RegistryKey? packagesKey = Registry.CurrentUser.OpenSubKey(
                @"Software\Classes\Extensions\ContractId\Windows.Launch\PackageId");
            if (packagesKey != null)
            {
                foreach (string subKeyName in packagesKey.GetSubKeyNames())
                {
                    if (subKeyName.Contains(NewOutlookPackageFamilyName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
        }
        catch
        {
            // Ignore: fall through to the App Paths check below.
        }

        try
        {
            using RegistryKey? appPathsKey = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\olk.exe");
            return appPathsKey != null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Returns true if the current process is running elevated (as Administrator). Used to
    /// distinguish "Outlook not running" from "Outlook is running, but at a different integrity
    /// level than this elevated process" (<c>GetActiveObject</c> fails with
    /// <c>MK_E_UNAVAILABLE</c> in that case).
    /// </summary>
    public static bool IsCurrentProcessElevated()
    {
        if (!OperatingSystem.IsWindows())
        {
            return false;
        }

        try
        {
            using WindowsIdentity identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Builds an actionable message describing why the classic Outlook COM object model is
    /// unavailable, distinguishing "not installed" from "new Outlook only" (unsupported) so the
    /// user does not have to guess. Use when <c>Type.GetTypeFromProgID("Outlook.Application")</c>
    /// returns null.
    /// </summary>
    public static string BuildUnavailableMessage(OutlookFlavor flavor) => flavor switch
    {
        OutlookFlavor.NewOutlookOnly =>
            "Only the new Outlook for Windows was detected, which has no COM object model and cannot be automated by this server. " +
            "Install or switch to classic Outlook for Windows (the desktop app with the 'Outlook.Application' COM ProgID) to use this tool.",
        OutlookFlavor.NotInstalled =>
            "Microsoft Outlook does not appear to be installed on this system. This server requires classic Outlook for Windows.",
        OutlookFlavor.Unknown =>
            "Could not determine which Outlook flavour, if any, is installed on this system. This server requires classic Outlook for Windows.",
        _ =>
            "Microsoft Outlook is not installed on this system."
    };
}
