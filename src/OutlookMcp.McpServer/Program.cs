using System.IO.Pipelines;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace OutlookMcp.McpServer;

/// <summary>
/// OutlookMcp Model Context Protocol (MCP) Server.
/// Hosts the Outlook tool surface (application, attachment, calendar, folder, mail).
/// </summary>
public class Program
{
    // Test transport configuration - set by tests before calling Main()
    // These are intentionally static for test injection. Thread-safety is not required
    // because tests run sequentially and call ResetTestTransport() after each test.
    private static Pipe? _testInputPipe;
    private static Pipe? _testOutputPipe;

    /// <summary>
    /// Configures the server to use in-memory pipe transport for testing.
    /// Call this before RunAsync() to enable test mode.
    /// </summary>
    /// <param name="inputPipe">Pipe for reading client requests (client writes, server reads)</param>
    /// <param name="outputPipe">Pipe for writing server responses (server writes, client reads)</param>
    public static void ConfigureTestTransport(Pipe inputPipe, Pipe outputPipe)
    {
        _testInputPipe = inputPipe;
        _testOutputPipe = outputPipe;
    }

    /// <summary>
    /// Resets test transport configuration (call after test completes).
    /// </summary>
    public static void ResetTestTransport()
    {
        _testInputPipe = null;
        _testOutputPipe = null;
    }

    /// <summary>
    /// The exact set of generated MCP tool types registered with the server (#23). This is an
    /// explicit allow-list rather than an assembly scan, so that a newly generated tool has to be
    /// added here on purpose instead of silently appearing and spending LLM context.
    ///
    /// The cost of that choice is that a genuinely new tool can be *forgotten* here and then be
    /// missing from tools/list while the CLI exposes it perfectly - which is exactly what happened
    /// to `contact`. <c>RegisteredToolsMatchGeneratedTools</c> exists to make that impossible to
    /// ship again, so this list must stay in sync with the generated <c>[McpServerToolType]</c>
    /// classes in this assembly.
    ///
    /// Declared as <see cref="IEnumerable{T}"/> deliberately. <c>WithTools</c> also has a generic
    /// <c>WithTools&lt;TToolType&gt;(builder, TToolType instance, ...)</c> overload, and passing a
    /// <c>Type[]</c> or <c>IReadOnlyList&lt;Type&gt;</c> binds to *that* by exact generic inference
    /// - registering the list object itself as a single tool instance and leaving the server with
    /// no tools at all, with no compiler error and no startup error. tools/list then answers
    /// "Method 'tools/list' is not available". Typing this as IEnumerable&lt;Type&gt; makes the
    /// non-generic overload win the tie-break.
    /// </summary>
    public static IEnumerable<Type> RegisteredToolTypes { get; } = new[]
    {
        typeof(OutlookMcp.McpServer.Tools.OutlookAddressbookTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookApplicationTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookAttachmentTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookCalendarTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookContactTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookFolderTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookMailTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookPropertyTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookRuleTool),
        typeof(OutlookMcp.McpServer.Tools.OutlookTaskTool),
    };

    public static async Task<int> Main(string[] args)
    {
        // Register assembly resolver for office.dll (Microsoft.Office.Core), which is a
        // .NET Framework GAC assembly that .NET Core cannot find via standard probing.
        // office.dll is copied to our output directory by Directory.Build.targets.
        RegisterOfficeAssemblyResolver();

        // Handle --help and --version flags for easy verification
        if (args.Length > 0)
        {
            var arg = args[0].ToLowerInvariant();
            if (arg is "-h" or "--help" or "-?" or "/?" or "/h")
            {
                ShowHelp();
                return 0;
            }
            if (arg is "-v" or "--version")
            {
                await ShowVersionAsync();
                return 0;
            }
        }

        // Register global exception handlers for unhandled exceptions
        RegisterGlobalExceptionHandlers();

        var builder = Host.CreateApplicationBuilder(args);

        // Disable FileSystemWatcher for config file reload.
        // Host.CreateApplicationBuilder() enables reloadOnChange:true by default, creating a
        // FileSystemWatcher for appsettings.json. Under file I/O storms (legacy presentation temp files, lock
        // files), this watcher fires ParseEventBufferAndNotifyForEach in a tight loop on the
        // threadpool, consuming ~85% CPU. Since MCP server config never changes at runtime,
        // disable reload entirely to eliminate the watcher.
        // Re-add JSON, environment variables, and CLI args — minus the file watchers.
        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine(args);

        // For stdio transport: Clear console logging to avoid polluting stderr with info messages.
        // The MCP client interprets stderr output as errors/warnings, so we only log Warning+
        // to stderr for debugging purposes. The MCP SDK handles protocol-level logging.
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            // Only log Warning and above to stderr - Info/Debug would appear as errors in MCP clients
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Warning;
        });
        builder.Logging.SetMinimumLevel(LogLevel.Warning);

        // Configure MCP Server - use test transport if configured, otherwise stdio
        var mcpBuilder = builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "outlook-mcp",
                    Version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0"
                };

                // Server-wide instructions for LLMs - helps with tool selection and workflow understanding
                options.ServerInstructions = """
                    OutlookMcp automates classic Outlook for Windows desktop via COM. It exposes
                    five tools: application, folder, mail, attachment, and calendar.

                    Identity model: Outlook items are addressed by entryId (optionally paired with
                    storeId for multi-store mailboxes), not by a file path or session handle. There
                    is no "open" or "close" step — pass entryId/storeId directly to the action that
                    needs the item.

                    Use useActiveMail/useActiveAppointment (a "read-active" style action) to operate
                    on whatever mail item or appointment the user currently has selected or open in
                    Outlook, when no explicit entryId is available. Prefer explicit entryId/storeId
                    targeting whenever you already have it, since "active item" is ambiguous with
                    nothing selected or multiple items shown.

                    Destructive actions (deleting items, sending mail, etc.) require confirmation
                    per the tool's Destructive annotation — do not chain a destructive action
                    immediately after a read without the user's explicit go-ahead, and never assume
                    a destructive call can be safely retried without checking whether it already
                    succeeded.
                    """;
            })
            // #23: explicit allow-list of Outlook-only types via the reflection-based WithTools/
            // WithPrompts(IEnumerable<Type>) overloads, instead of .WithToolsFromAssembly()/
            // .WithPromptsFromAssembly(). Registering by an explicit list keeps tools/list
            // deterministic: anything newly generated has to be added here on purpose rather than
            // silently appearing and wasting LLM context. The generic WithTools<T>/WithPrompts<T>
            // overloads require a non-static T, but the generated tool/prompt classes are static,
            // so the non-generic Type-list overloads are used instead.
            // The list lives on RegisteredToolTypes so a test can assert it covers every generated
            // tool; see the comment there.
            // OutlookSkillPrompts.g.cs is itself narrowed to only Outlook-relevant prompts (see
            // GenerateSkillPrompts target in the .csproj), so registering the whole type is safe.
            .WithTools(RegisteredToolTypes)
            .WithPrompts([typeof(OutlookMcp.McpServer.Prompts.OutlookSkillPrompts)]);

        if (_testInputPipe != null && _testOutputPipe != null)
        {
            // Test mode: use in-memory pipe transport
            mcpBuilder.WithStreamServerTransport(
                _testInputPipe.Reader.AsStream(),
                _testOutputPipe.Writer.AsStream());
        }
        else
        {
            // Production mode: use stdio transport
            mcpBuilder.WithStdioServerTransport();
        }

        var host = builder.Build();

        // Note: Update checks are handled by the service layer (shown via Windows notification)
        // to avoid duplicate notifications when running in unified package mode

        try
        {
            await host.RunAsync();
            return 0;
        }
        catch (OperationCanceledException)
        {
            // Graceful shutdown via cancellation (e.g., Ctrl+C, SIGTERM)
            // This is expected behavior, not an error
            return 0;
        }
#pragma warning disable CA1031 // Catch general exception - this is a top-level handler that must not crash
        catch (Exception ex)
        {
            // Return exit code 1 for fatal errors (FR-024, SC-015a)
            // Do NOT re-throw - deterministic exit code is more important for callers
            Console.Error.WriteLine($"[OutlookMcp] Fatal error: {ex.Message}");
            return 1;
        }
#pragma warning restore CA1031
        finally
        {
            // CRITICAL: Auto-save legacy presentation sessions on shutdown.
            // Without this, MCP client disconnect or process exit silently discards unsaved legacy session work.
            ServiceBridge.ServiceBridge.Dispose();
        }
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        // Handle exceptions that escape all catch blocks
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                Console.Error.WriteLine($"[OutlookMcp] Unhandled exception: {ex.Message}");
            }
        };

        // Handle unobserved task exceptions
        TaskScheduler.UnobservedTaskException += (sender, e) =>
        {
            Console.Error.WriteLine($"[OutlookMcp] Unobserved task exception: {e.Exception.Message}");
        };
    }

    /// <summary>
    /// Registers assembly resolver for office.dll (Microsoft.Office.Core).
    /// </summary>
    private static void RegisterOfficeAssemblyResolver()
    {
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name);
            if (!string.Equals(name.Name, "office", StringComparison.OrdinalIgnoreCase))
                return null;

            return ResolveOfficeDll();
        };
    }

    /// <summary>
    /// Resolves office.dll (Microsoft.Office.Core) from multiple locations.
    /// office.dll is a .NET Framework GAC assembly that .NET Core cannot find automatically.
    /// It is present when Microsoft Office is installed, but not in the .NET Core probing paths.
    /// Search order:
    ///   1. AppContext.BaseDirectory (copied by Directory.Build.targets in local dev builds)
    ///   2. .NET Framework GAC - v16 then v15 (v15 is accepted by the CLR for v16 requests)
    ///   3. Office installation directory (click-to-run Office 365 doesn't register in GAC)
    /// </summary>
    private static Assembly? ResolveOfficeDll()
    {
        // 1. Local build output (Directory.Build.targets copies office.dll here in dev builds)
        var localPath = Path.Combine(AppContext.BaseDirectory, "office.dll");
        if (File.Exists(localPath))
            return Assembly.LoadFrom(localPath);

        // 2. .NET Framework GAC — v16 preferred, v15 accepted (CLR honours AssemblyResolve return regardless of version)
        string[] gacPaths =
        [
            @"C:\Windows\assembly\GAC_MSIL\office\16.0.0.0__71e9bce111e9429c\OFFICE.DLL",
            @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL",
        ];
        foreach (var gacPath in gacPaths)
        {
            if (File.Exists(gacPath))
                return Assembly.LoadFrom(gacPath);
        }

        // 3. Office 365 click-to-run installation directories (Office registers its own copy)
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string[] officeDirs =
        [
            Path.Combine(programFiles, @"Microsoft Office\root\Office16"),
            Path.Combine(programFilesX86, @"Microsoft Office\root\Office16"),
        ];
        foreach (var dir in officeDirs)
        {
            var officePath = Path.Combine(dir, "OFFICE.dll");
            if (File.Exists(officePath))
                return Assembly.LoadFrom(officePath);
        }

        return null;
    }

    /// <summary>
    /// Shows help information.
    /// </summary>
    private static void ShowHelp()
    {
        var version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0";
        Console.WriteLine($"""
            Outlook MCP Server (Migration) v{version}

            An MCP (Model Context Protocol) server for the Outlook migration surface.

            Usage:
              OutlookMcp.McpServer.exe [options]

            Options:
              -h, --help      Show this help message
              -v, --version   Show version information

            Without options, starts the MCP server in stdio mode.

            Requirements:
              - Windows x64
              - Classic Microsoft Outlook desktop installed
            """);
    }

    /// <summary>
    /// Shows version information and checks for updates.
    /// </summary>
    private static async Task ShowVersionAsync()
    {
        var currentVersion = Infrastructure.McpServerVersionChecker.GetCurrentVersion();
        Console.WriteLine($"Outlook MCP Server (Migration) v{currentVersion}");

        // Check for updates (non-blocking, 5-second timeout)
        var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();
        if (latestVersion != null)
        {
            Console.WriteLine();
            Console.WriteLine($"Update available: {currentVersion} -> {latestVersion}");
            Console.WriteLine("Run: dotnet tool update --global OutlookMcp.McpServer");
            Console.WriteLine("Release notes: https://github.com/trsdn/mcp-server-outlook/releases/latest");
        }
    }
}



