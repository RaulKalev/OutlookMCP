using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Infrastructure.Exchange;

/// <summary>
/// Owns the MSAL public-client application for the configured Entra app registration.
/// Sign-in uses the OAuth device-code flow; tokens are persisted in an encrypted cache
/// where the platform supports it, so the user authorizes once and later calls refresh
/// silently.
/// </summary>
public sealed class ExchangeTokenProvider(OutlookMcpOptions options, ILogger<ExchangeTokenProvider> logger) : IDisposable
{
    private static readonly string[] Scopes = ["Calendars.ReadWrite"];

    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly CalendarSyncOptions _options = options.CalendarSync;
    private IPublicClientApplication? _app;
    private bool _cachePersisted;
    private Task<AuthenticationResult>? _pendingLogin;
    private DeviceCodeResult? _pendingCode;

    public async Task<ExchangeAuthStatusDto> GetStatusAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var app = await GetAppAsync().ConfigureAwait(false);
            CollectCompletedLogin();
            if (_pendingLogin is { IsCompleted: false } && _pendingCode is not null)
            {
                return new ExchangeAuthStatusDto("login_pending", null, _cachePersisted,
                    $"Sign-in is waiting for the user: open {_pendingCode.VerificationUrl} and enter code {_pendingCode.UserCode}.");
            }

            var account = (await app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
            return account is null
                ? new ExchangeAuthStatusDto("signed_out", null, _cachePersisted, "No Exchange account is signed in. Run outlook_exchange_login.")
                : new ExchangeAuthStatusDto("signed_in", account.Username, _cachePersisted, "Signed in; tokens refresh silently.");
        }
        finally { _gate.Release(); }
    }

    public async Task<ExchangeLoginDto> BeginDeviceCodeLoginAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var app = await GetAppAsync().ConfigureAwait(false);
            CollectCompletedLogin();
            var account = (await app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
            if (account is not null)
            {
                return new ExchangeLoginDto("already_signed_in", account.Username, null, null, null,
                    "An Exchange account is already signed in. Run outlook_exchange_logout first to switch accounts.");
            }

            if (_pendingLogin is { IsCompleted: false } && _pendingCode is not null)
            {
                return PendingLoginDto();
            }

            var codeReady = new TaskCompletionSource<DeviceCodeResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            var login = app.AcquireTokenWithDeviceCode(Scopes, deviceCode =>
            {
                codeReady.TrySetResult(deviceCode);
                return Task.CompletedTask;
            }).ExecuteAsync(CancellationToken.None);
            var completed = await Task.WhenAny(codeReady.Task, login).ConfigureAwait(false);
            if (ReferenceEquals(completed, login) && !codeReady.Task.IsCompleted)
            {
                await ObserveFailedLoginAsync(login).ConfigureAwait(false);
            }

            _pendingLogin = login;
            _pendingCode = await codeReady.Task.ConfigureAwait(false);
            _ = login.ContinueWith(
                value => logger.LogInformation("Exchange device-code login finished with status {Status}", value.Status),
                CancellationToken.None, TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
            return PendingLoginDto();
        }
        finally { _gate.Release(); }
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var app = await GetAppAsync().ConfigureAwait(false);
            CollectCompletedLogin();
            var account = (await app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault()
                ?? throw new OutlookMcpException(ErrorCodes.ExchangeAuthRequired, "No Exchange account is signed in.",
                    "Run outlook_exchange_login, complete the device-code sign-in, then retry.");
            try
            {
                var result = await app.AcquireTokenSilent(Scopes, account).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                return result.AccessToken;
            }
            catch (MsalUiRequiredException ex)
            {
                throw new OutlookMcpException(ErrorCodes.ExchangeAuthRequired, "The Exchange sign-in has expired or requires new consent.",
                    "Run outlook_exchange_login again to refresh the sign-in.", ex);
            }
            catch (MsalException ex)
            {
                throw new OutlookMcpException(ErrorCodes.ExchangeApiFailed, "Microsoft sign-in could not issue an access token.",
                    "Retry; if the problem continues, run outlook_exchange_login again.", ex);
            }
        }
        finally { _gate.Release(); }
    }

    public async Task<ExchangeAuthStatusDto> LogoutAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var app = await GetAppAsync().ConfigureAwait(false);
            foreach (var account in await app.GetAccountsAsync().ConfigureAwait(false))
            {
                await app.RemoveAsync(account).ConfigureAwait(false);
            }

            _pendingLogin = null;
            _pendingCode = null;
            return new ExchangeAuthStatusDto("signed_out", null, _cachePersisted, "Signed out. Cached tokens for this app were removed.");
        }
        finally { _gate.Release(); }
    }

    private ExchangeLoginDto PendingLoginDto() => new(
        "pending", null, _pendingCode!.VerificationUrl, _pendingCode.UserCode, _pendingCode.ExpiresOn,
        $"Ask the user to open {_pendingCode.VerificationUrl} in any browser, enter code {_pendingCode.UserCode}, and approve the Calendars.ReadWrite permission. Check outlook_exchange_auth_status afterwards.");

    private void CollectCompletedLogin()
    {
        if (_pendingLogin is null || !_pendingLogin.IsCompleted) return;
        if (_pendingLogin.IsFaulted)
        {
            logger.LogWarning(_pendingLogin.Exception?.GetBaseException(), "The pending Exchange device-code login failed");
        }

        _pendingLogin = null;
        _pendingCode = null;
    }

    private static async Task ObserveFailedLoginAsync(Task<AuthenticationResult> login)
    {
        try
        {
            await login.ConfigureAwait(false);
            throw new OutlookMcpException(ErrorCodes.ExchangeApiFailed, "The device-code login ended before producing a sign-in code.",
                "Retry outlook_exchange_login.");
        }
        catch (MsalException ex)
        {
            throw new OutlookMcpException(ErrorCodes.ExchangeApiFailed, "Microsoft rejected the device-code sign-in request.",
                "Verify CalendarSync.ClientId and TenantId in config.json, and that the app registration has 'Allow public client flows' enabled.", ex);
        }
    }

    public void Dispose() => _gate.Dispose();

    private async Task<IPublicClientApplication> GetAppAsync()
    {
        if (_app is not null) return _app;
        var clientId = _options.ClientId;
        if (string.IsNullOrWhiteSpace(clientId))
        {
            throw new OutlookMcpException(ErrorCodes.ExchangeNotConfigured, "CalendarSync.ClientId is not configured.",
                "Register a free Entra ID app (see the README's calendar sync section) and set CalendarSync.ClientId in config.json.");
        }

        var app = PublicClientApplicationBuilder.Create(clientId.Trim())
            .WithAuthority($"https://login.microsoftonline.com/{_options.TenantId.Trim()}")
            .Build();
        try
        {
            var directory = AppPaths.ExpandPath(_options.TokenCacheDirectory);
            Directory.CreateDirectory(directory);
            var storage = new StorageCreationPropertiesBuilder("msal-token-cache.bin", directory).Build();
            var cacheHelper = await MsalCacheHelper.CreateAsync(storage).ConfigureAwait(false);
            cacheHelper.VerifyPersistence();
            cacheHelper.RegisterCache(app.UserTokenCache);
            _cachePersisted = true;
        }
        catch (MsalCachePersistenceException ex)
        {
            _cachePersisted = false;
            logger.LogWarning(ex, "The encrypted token cache is unavailable; Exchange sign-in will not survive a server restart");
        }

        _app = app;
        return app;
    }
}
