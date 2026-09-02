using System.Reflection;
using System.Runtime.InteropServices;

namespace OutlookMcp.CLI.Infrastructure;

/// <summary>
/// System tray icon for the CLI daemon process. Shows service status and allows stopping it.
/// </summary>
internal sealed class CliServiceTray : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly ContextMenuStrip _contextMenu;
    private readonly Action _requestShutdown;
    private readonly TaskbarNotificationWindow _taskbarWindow;
    private bool _disposed;

    public CliServiceTray(Action requestShutdown)
    {
        _requestShutdown = requestShutdown;

        _contextMenu = new ContextMenuStrip();

        // About (Alt+A mnemonic)
        var aboutItem = new ToolStripMenuItem("&About...");
        aboutItem.AccessibleDescription = "Show version and project information";
        aboutItem.Click += (_, _) => ShowAbout();
        _contextMenu.Items.Add(aboutItem);

        _contextMenu.Items.Add(new ToolStripSeparator());

        // Exit (Alt+X mnemonic)
        var exitItem = new ToolStripMenuItem("E&xit");
        exitItem.AccessibleDescription = "Stop the CLI service and exit";
        exitItem.Click += (_, _) => ExitService();
        _contextMenu.Items.Add(exitItem);

        // Load icon
        var icon = LoadEmbeddedIcon();

        _notifyIcon = new NotifyIcon
        {
            Icon = icon,
            Text = "OutlookMcp CLI Service",
            ContextMenuStrip = _contextMenu,
            Visible = true
        };

        // Listen for explorer.exe restarts so we can re-register the tray icon
        _taskbarWindow = new TaskbarNotificationWindow(_notifyIcon);

        // Check for updates after a short delay so the UI is responsive at startup
        CheckForUpdateAsync();
    }

    private static Icon LoadEmbeddedIcon()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = "OutlookMcp.CLI.Resources.outlookcli.ico";

        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream != null)
        {
            return new Icon(stream);
        }

        return SystemIcons.Application;
    }

    /// <summary>
    /// Checks NuGet for a newer version after a 5-second delay and shows a balloon tip if available.
    /// </summary>
    private async void CheckForUpdateAsync()
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(5));

            if (_disposed) return;

            var currentVersion = GetCurrentVersion();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cts.Token);

            if (_disposed || latestVersion == null) return;

            if (CompareVersions(currentVersion, latestVersion) < 0)
            {
                ShowBalloon(
                    "Update Available",
                    $"OutlookMcp CLI {latestVersion} is available (current: {currentVersion}).\n" +
                    "Run: dotnet tool update --global OutlookMcp.CLI");
            }
        }
        catch
        {
            // Version check should never crash the service
        }
    }

    private void ShowBalloon(string title, string message, ToolTipIcon icon = ToolTipIcon.Info)
    {
        _notifyIcon.ShowBalloonTip(3000, title, message, icon);
    }

    private static async void ShowAbout()
    {
        var version = GetCurrentVersion();

        string? latestVersion = null;
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cts.Token);
        }
        catch
        {
            // Version check failed — show dialog without update info
        }

        var updateAvailable = latestVersion != null && CompareVersions(version, latestVersion) < 0;

        using var form = new Form
        {
            Text = "About Outlook Migration CLI",
            Size = new Size(420, updateAvailable ? 300 : 260),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            KeyPreview = true
        };

        // Allow Escape to close the dialog
        form.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape) form.Close();
        };

        var iconBox = new PictureBox
        {
            Image = SystemIcons.Information.ToBitmap(),
            SizeMode = PictureBoxSizeMode.AutoSize,
            Location = new Point(20, 20),
            AccessibleName = "Outlook Migration CLI icon",
            AccessibleRole = AccessibleRole.Graphic,
            TabStop = false
        };

        var nameLabel = new Label
        {
            Text = "Outlook Migration CLI Service",
            Font = new Font(Control.DefaultFont.FontFamily, 10, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(70, 20),
            AccessibleName = "Outlook Migration CLI Service",
            AccessibleRole = AccessibleRole.StaticText
        };

        var versionLabel = new Label
        {
            Text = $"Version: {version}",
            AutoSize = true,
            Location = new Point(70, 45),
            AccessibleName = $"Version {version}",
            AccessibleRole = AccessibleRole.StaticText
        };

        var descLabel = new Label
        {
            Text = "Outlook automation service.",
            AutoSize = true,
            Location = new Point(70, 75),
            AccessibleName = "Outlook automation service",
            AccessibleRole = AccessibleRole.StaticText
        };

        const string githubUrl = "https://github.com/trsdn/mcp-server-outlook";
        const string docsUrl = "https://github.com/trsdn/mcp-server-outlook";

        var githubLabel = new Label
        {
            Text = "GitHub:",
            AutoSize = true,
            Location = new Point(70, 105),
            AccessibleRole = AccessibleRole.StaticText
        };
        var githubLink = new LinkLabel
        {
            Text = githubUrl,
            AutoSize = true,
            Location = new Point(125, 105),
            TabIndex = 0,
            AccessibleName = "GitHub repository link",
            AccessibleDescription = $"Opens {githubUrl} in browser"
        };
        githubLink.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(githubUrl) { UseShellExecute = true }); }
            catch { /* Ignore navigation errors */ }
        };

        var docsLabel = new Label
        {
            Text = "Docs:",
            AutoSize = true,
            Location = new Point(70, 130),
            AccessibleRole = AccessibleRole.StaticText
        };
        var docsLink = new LinkLabel
        {
            Text = docsUrl,
            AutoSize = true,
            Location = new Point(125, 130),
            TabIndex = 1,
            AccessibleName = "Documentation link",
            AccessibleDescription = $"Opens {docsUrl} in browser"
        };
        docsLink.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docsUrl) { UseShellExecute = true }); }
            catch { /* Ignore navigation errors */ }
        };

        var tabIndex = 2;
        var buttonY = 165;
        form.Controls.AddRange([iconBox, nameLabel, versionLabel, descLabel, githubLabel, githubLink, docsLabel, docsLink]);

        if (updateAvailable)
        {
            var updateLabel = new Label
            {
                Text = $"Update available: {version} \u2192 {latestVersion}",
                ForeColor = SystemColors.HotTrack,
                Font = new Font(Control.DefaultFont, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(70, 160),
                AccessibleName = $"Update available from version {version} to {latestVersion}",
                AccessibleRole = AccessibleRole.StaticText
            };

            var updateCmd = new TextBox
            {
                Text = "dotnet tool update --global OutlookMcp.CLI",
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = form.BackColor,
                Location = new Point(70, 180),
                Size = new Size(320, 20),
                TabIndex = tabIndex++,
                AccessibleName = "Update command, select to copy",
                AccessibleDescription = "Run this command in a terminal to update"
            };

            form.Controls.AddRange([updateLabel, updateCmd]);
            buttonY = 210;
        }

        var okButton = new Button
        {
            Text = "&OK",
            DialogResult = System.Windows.Forms.DialogResult.OK,
            Size = new Size(80, 28),
            Location = new Point(160, buttonY),
            TabIndex = tabIndex,
            AccessibleName = "OK, close dialog"
        };
        form.AcceptButton = okButton;
        form.Controls.Add(okButton);

        form.ShowDialog();
    }

    private static int CompareVersions(string current, string latest)
    {
        if (Version.TryParse(current, out var currentVer) && Version.TryParse(latest, out var latestVer))
            return currentVer.CompareTo(latestVer);
        return string.Compare(current, latest, StringComparison.Ordinal);
    }

    private static string GetCurrentVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        return informational?.Split('+')[0] ?? assembly.GetName().Version?.ToString() ?? "0.0.0";
    }


    private void ExitService()
    {
        _requestShutdown();
    }
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _taskbarWindow.DestroyHandle();
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _contextMenu.Dispose();
    }

    /// <summary>
    /// Hidden window that listens for the TaskbarCreated message broadcast by explorer.exe
    /// after it restarts, so the tray icon can be re-registered.
    /// </summary>
    private sealed class TaskbarNotificationWindow : NativeWindow
    {
        private readonly NotifyIcon _notifyIcon;
        private readonly uint _wmTaskbarCreated;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpString);

        public TaskbarNotificationWindow(NotifyIcon notifyIcon)
        {
            _notifyIcon = notifyIcon;
            _wmTaskbarCreated = RegisterWindowMessage("TaskbarCreated");

            // Create a message-only window to receive broadcast messages
            CreateHandle(new CreateParams());
        }

        protected override void WndProc(ref Message m)
        {
            if (_wmTaskbarCreated != 0 && m.Msg == (int)_wmTaskbarCreated)
            {
                // Explorer restarted — re-register the tray icon
                _notifyIcon.Visible = false;
                _notifyIcon.Visible = true;
            }

            base.WndProc(ref m);
        }
    }

}
