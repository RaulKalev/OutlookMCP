using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Services;
using OutlookMcp.OutlookInterop;
using OutlookMcp.Server.Mcp;
using Serilog;
using Serilog.Events;
using Serilog.Formatting.Compact;

return await OutlookMcpProgram.RunAsync(args).ConfigureAwait(false);

internal static class OutlookMcpProgram
{
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
        Directory.CreateDirectory(AppPaths.LogDirectory);
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Is(ParseLevel(options.Logging.Level))
            .Enrich.FromLogContext()
            .WriteTo.File(new CompactJsonFormatter(), Path.Combine(AppPaths.LogDirectory, "outlook-mcp-.jsonl"), rollingInterval: RollingInterval.Day, retainedFileCountLimit: Math.Clamp(options.Logging.RetentionDays, 1, 365), fileSizeLimitBytes: 20 * 1024 * 1024, rollOnFileSizeLimit: true, shared: false)
            .CreateLogger();

        try
        {
            using var host = BuildHost(args, options);
            if (args.Contains("--diagnose", StringComparer.OrdinalIgnoreCase)) return await DiagnoseAsync(host).ConfigureAwait(false);
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

    private static IHost BuildHost(string[] args, OutlookMcpOptions options)
    {
        var builder = Host.CreateApplicationBuilder(args);
        builder.Logging.ClearProviders();
        builder.Services.AddSerilog();
        builder.Services.AddSingleton(options);
        builder.Services.AddSingleton<EmailBodyCleaner>();
        builder.Services.AddSingleton<OutlookStaDispatcher>();
        builder.Services.AddSingleton<OutlookGateway>();
        builder.Services.AddSingleton<IOutlookGateway>(provider => provider.GetRequiredService<OutlookGateway>());
        builder.Services.AddSingleton<ToolExecutor>();
        builder.Services.AddMcpServer().WithStdioServerTransport().WithToolsFromAssembly();
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

    private static LogEventLevel ParseLevel(string value) => Enum.TryParse<LogEventLevel>(value, true, out var level) ? level : LogEventLevel.Information;
}
