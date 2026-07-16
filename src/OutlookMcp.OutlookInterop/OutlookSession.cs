using System.Diagnostics;
using System.Runtime.InteropServices;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.OutlookInterop;

internal sealed class OutlookSession : IDisposable
{
    private readonly OutlookOptions _options;
    private dynamic? _application;
    private dynamic? _namespace;
    private bool _startedByServer;

    public OutlookSession(OutlookOptions options) => _options = options;

    public dynamic Application => _application ?? throw NotAvailable();
    public dynamic Namespace => _namespace ?? throw new OutlookMcpException(ErrorCodes.MapiNotReady, "The Outlook MAPI namespace is not ready.", "Open Outlook Classic, finish profile setup, and retry.");
    public bool WasRunningAtConnection { get; private set; }

    public bool IsInstalled => Type.GetTypeFromProgID("Outlook.Application", throwOnError: false) is not null;

    public void EnsureConnected()
    {
        if (_application is not null && _namespace is not null) return;
        var outlookType = Type.GetTypeFromProgID("Outlook.Application", throwOnError: false)
            ?? throw new OutlookMcpException(ErrorCodes.OutlookNotInstalled, "Microsoft Outlook Classic is not installed or its COM registration is unavailable.", "Install or repair Outlook Classic for Windows. New Outlook is not supported.");

        WasRunningAtConnection = Process.GetProcessesByName("OUTLOOK").Length > 0;
        if (!WasRunningAtConnection && !_options.StartIfNotRunning)
        {
            throw new OutlookMcpException(ErrorCodes.OutlookNotAvailable, "Outlook Classic is not running and automatic startup is disabled.", "Start Outlook Classic or enable Outlook.StartIfNotRunning.");
        }

        try
        {
            _application = Activator.CreateInstance(outlookType) ?? throw new InvalidOperationException("COM activation returned null.");
            _startedByServer = !WasRunningAtConnection;
            _namespace = _application.GetNamespace("MAPI");
            var deadline = DateTime.UtcNow.AddSeconds(Math.Clamp(_options.OperationTimeoutSeconds, 1, 300));
            while (DateTime.UtcNow < deadline)
            {
                dynamic? stores = null;
                try
                {
                    stores = _namespace.Stores;
                    _ = (int)stores.Count;
                    return;
                }
                catch (COMException) { Thread.Sleep(250); }
                finally { ComReleaseHelper.FinalRelease(stores); }
            }

            throw new OutlookMcpException(ErrorCodes.MapiNotReady, "Outlook started, but its profile and stores did not become ready in time.", "Open Outlook Classic, complete any profile prompts, then retry.");
        }
        catch (OutlookMcpException) { throw; }
        catch (Exception ex) when (ex is COMException or InvalidCastException or InvalidOperationException)
        {
            ReleaseSession();
            throw new OutlookMcpException(ErrorCodes.OutlookNotAvailable, "Outlook Classic could not be activated.", "Start Outlook Classic manually and verify that a mail profile opens successfully.", ex);
        }
    }

    public void Dispose()
    {
        if (_startedByServer && _application is not null)
        {
            try { _application.Quit(); }
            catch (COMException) { }
        }

        ReleaseSession();
    }

    private void ReleaseSession()
    {
        ComReleaseHelper.FinalRelease(_namespace);
        ComReleaseHelper.FinalRelease(_application);
        _namespace = null;
        _application = null;
    }

    private static OutlookMcpException NotAvailable() => new(ErrorCodes.OutlookNotAvailable, "Outlook Classic is unavailable.", "Start Outlook Classic and retry.");
}
