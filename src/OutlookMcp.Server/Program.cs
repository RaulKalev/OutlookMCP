using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Services;
using OutlookMcp.Application.WritingStyle;
using OutlookMcp.Infrastructure.Exchange;
using OutlookMcp.Infrastructure.WritingStyle;
using OutlookMcp.OutlookInterop;
using OutlookMcp.Server.Mcp;
using Serilog;
using Serilog.Events;
using Serilog.Formatting.Compact;

return await OutlookMcpProgram.RunAsync(args).ConfigureAwait(false);

internal static class OutlookMcpProgram
{
    private static readonly string[] StyleMaintenanceCommands = ["--style-scan-status", "--style-scan", "--style-sync", "--style-rebuild-profile"];

    public static async Task<int> RunAsync(string[] args)
    {
        if (args.Contains("--version", StringComparer.OrdinalIgnoreCase))
        {
            Console.WriteLine(Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0");
            return 0;
        }
        if (args.Contains("--print-config-path", StringComparer.OrdinalIgnoreCase)) { Console.WriteLine(AppPaths.ConfigPath); return 0; }
        if (args.Contains("--print-log-path", StringComparer.OrdinalIgnoreCase)) { Console.WriteLine(AppPaths.LogDirectory); return 0; }

        EnsureUserConfiguration();
        var options = LoadOptions();
        ApplyCommandLinePaths(args, options);
        var toolProfile = McpToolProfileParser.Parse(args);
        OutlookMcpOptionsValidator.Validate(options);
        Directory.CreateDirectory(AppPaths.LogDirectory);
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Is(ParseLevel(options.Logging.Level))
            .Enrich.FromLogContext()
            .WriteTo.File(new CompactJsonFormatter(), Path.Combine(AppPaths.LogDirectory, "outlook-mcp-.jsonl"), rollingInterval: RollingInterval.Day, retainedFileCountLimit: Math.Clamp(options.Logging.RetentionDays, 1, 365), fileSizeLimitBytes: 20 * 1024 * 1024, rollOnFileSizeLimit: true, shared: false)
            .CreateLogger();

        try
        {
            using var host = BuildHost(args, options, toolProfile);
            if (args.Contains("--diagnose", StringComparer.OrdinalIgnoreCase)) return await DiagnoseAsync(host).ConfigureAwait(false);
            var styleExitCode = await RunStyleMaintenanceAsync(host, args).ConfigureAwait(false);
            if (styleExitCode is not null) return styleExitCode.Value;
            await host.RunAsync().ConfigureAwait(false);
            return 0;
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Outlook MCP server terminated unexpectedly");
            if (args.Contains("--diagnose", StringComparer.OrdinalIgnoreCase)) Console.Error.WriteLine($"FAIL: {ex.Message}");
            return 1;
        }
        finally { await Log.CloseAndFlushAsync().ConfigureAwait(false); }
    }

    private static IHost BuildHost(string[] args, OutlookMcpOptions options, McpToolProfile toolProfile)
    {
        var builder = Host.CreateApplicationBuilder(args);
        builder.Logging.ClearProviders();
        builder.Services.AddSerilog();
        builder.Services.AddSingleton(options);
        builder.Services.AddSingleton<EmailBodyCleaner>();
        builder.Services.AddSingleton<OutlookStaDispatcher>();
        builder.Services.AddSingleton<OutlookGateway>();
        builder.Services.AddSingleton<IOutlookGateway>(provider => provider.GetRequiredService<OutlookGateway>());
        builder.Services.AddSingleton<AuthoredTextExtractor>();
        builder.Services.AddSingleton<CommunicationIntentClassifier>();
        builder.Services.AddSingleton<WritingProfileStore>();
        builder.Services.AddSingleton<SqliteStyleIndexRepository>();
        builder.Services.AddSingleton<IStyleIndexRepository>(provider => provider.GetRequiredService<SqliteStyleIndexRepository>());
        builder.Services.AddSingleton<WritingStyleCoordinator>();
        builder.Services.AddSingleton<ExchangeTokenProvider>();
        builder.Services.AddSingleton<GraphCalendarClient>();
        builder.Services.AddSingleton<IExchangeCalendarGateway>(provider => provider.GetRequiredService<GraphCalendarClient>());
        builder.Services.AddSingleton<CalendarSyncCoordinator>();
        builder.Services.AddSingleton<ToolExecutor>();
        var jsonOptions = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            TypeInfoResolver = new DefaultJsonTypeInfoResolver()
        };
        var mcp = builder.Services.AddMcpServer().WithStdioServerTransport();
        if (toolProfile is McpToolProfile.Compact)
        {
            mcp.WithTools<OutlookCompactTools>(jsonOptions);
        }
        else if (toolProfile is McpToolProfile.Mail)
        {
            mcp.WithTools<OutlookTools>(jsonOptions);
        }
        else if (toolProfile is McpToolProfile.Style)
        {
            mcp.WithTools<OutlookStyleTools>(jsonOptions);
        }
        else
        {
            mcp.WithTools<OutlookTools>(jsonOptions);
            mcp.WithTools<OutlookStyleTools>(jsonOptions);
        }
        return builder.Build();
    }

    private static OutlookMcpOptions LoadOptions()
    {
        var configuration = new ConfigurationBuilder().AddJsonFile(AppPaths.ConfigPath, optional: false, reloadOnChange: false).Build();
        return OutlookMcpOptionsValidator.Validate(configuration.Get<OutlookMcpOptions>() ?? new OutlookMcpOptions());
    }

    private static void EnsureUserConfiguration()
    {
        Directory.CreateDirectory(AppPaths.ConfigDirectory);
        if (File.Exists(AppPaths.ConfigPath)) return;
        var json = JsonSerializer.Serialize(new OutlookMcpOptions(), new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(AppPaths.ConfigPath, json);
    }

    private static async Task<int> DiagnoseAsync(IHost host)
    {
        Console.WriteLine("EULE Outlook MCP diagnostics (no email bodies are read or printed)");
        Console.WriteLine($"Configuration: OK ({AppPaths.ConfigPath})");
        Console.WriteLine($"Logs: OK ({AppPaths.LogDirectory})");
        var gateway = host.Services.GetRequiredService<IOutlookGateway>();
        try
        {
            var status = await gateway.GetStatusAsync(CancellationToken.None).ConfigureAwait(false);
            Console.WriteLine($"Outlook Classic installed: {status.OutlookClassicInstalled}");
            Console.WriteLine($"Outlook process available: {status.OutlookRunning}");
            Console.WriteLine($"MAPI namespace available: {status.MapiAvailable}");
            Console.WriteLine($"Profile: {status.ProfileName ?? "(unavailable)"}");
            Console.WriteLine($"Stores: {status.StoreCount}");
            var stores = await gateway.ListStoresAsync(CancellationToken.None).ConfigureAwait(false);
            Console.WriteLine($"Allowed accessible stores: {stores.Count(value => value.IsAccessible)}/{stores.Count}");
            var outlookDiagnostics = await host.Services.GetRequiredService<OutlookGateway>().DiagnoseOutlookAsync(CancellationToken.None).ConfigureAwait(false);
            var attachmentDirectory = new AttachmentPathPolicy(host.Services.GetRequiredService<OutlookMcpOptions>().Outlook).DefaultDirectory;
            Directory.CreateDirectory(attachmentDirectory);
            var testPath = Path.Combine(attachmentDirectory, $".outlook-mcp-write-test-{Guid.NewGuid():N}");
            await File.WriteAllTextAsync(testPath, string.Empty).ConfigureAwait(false);
            File.Delete(testPath);
            Console.WriteLine($"Attachment directory writable: True ({attachmentDirectory})");
            Console.WriteLine($"Draft folder accessible: {outlookDiagnostics.DraftFolderAccessible}");
            Console.WriteLine($"Explorer/Inspector access: {outlookDiagnostics.SelectedItemAccessible} (no item content inspected)");
            Console.WriteLine("RESULT: OK");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"RESULT: FAIL - {ex.Message}");
            return 2;
        }
    }

    private static async Task<int?> RunStyleMaintenanceAsync(IHost host, string[] args)
    {
        var command = StyleMaintenanceCommands
            .FirstOrDefault(value => args.Contains(value, StringComparer.OrdinalIgnoreCase));
        if (command is null) return null;
        var style = host.Services.GetRequiredService<WritingStyleCoordinator>();
        object result = command switch
        {
            "--style-scan-status" => await style.GetStatusAsync(CancellationToken.None).ConfigureAwait(false),
            "--style-scan" => await style.ScanAsync(null, 10_000, false, true, CancellationToken.None).ConfigureAwait(false),
            "--style-sync" => await style.SyncAsync(CancellationToken.None).ConfigureAwait(false),
            "--style-rebuild-profile" => await style.SaveBaselineProfileAsync("Generated by --style-rebuild-profile from local aggregate statistics.", CancellationToken.None).ConfigureAwait(false),
            _ => throw new InvalidOperationException("Unsupported style maintenance command.")
        };
        Console.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }

    private static void ApplyCommandLinePaths(string[] args, OutlookMcpOptions options)
    {
        var database = ArgumentValue(args, "--style-db-path");
        var profile = ArgumentValue(args, "--style-profile-path");
        if (!string.IsNullOrWhiteSpace(database)) options.WritingStyle.DatabasePath = database;
        if (!string.IsNullOrWhiteSpace(profile))
        {
            options.WritingStyle.ProfilePath = profile;
            var directory = Path.GetDirectoryName(Path.GetFullPath(profile)) ?? Environment.CurrentDirectory;
            options.WritingStyle.ProfileHistoryPath = Path.Combine(directory, Path.GetFileNameWithoutExtension(profile) + "-history.json");
        }
    }

    private static string? ArgumentValue(string[] args, string name)
    {
        var index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException($"{name} requires a path value.");
        return args[index + 1];
    }

    private static LogEventLevel ParseLevel(string value) => Enum.TryParse<LogEventLevel>(value, true, out var level) ? level : LogEventLevel.Information;
}
