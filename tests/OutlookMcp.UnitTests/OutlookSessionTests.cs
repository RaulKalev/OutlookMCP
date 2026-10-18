using System.Runtime.InteropServices;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;
using OutlookMcp.OutlookInterop;

namespace OutlookMcp.UnitTests;

public sealed class OutlookSessionTests
{
    [Theory]
    [InlineData(0x800706BA)] // RPC_S_SERVER_UNAVAILABLE: Outlook restarted after the server connected
    [InlineData(0x80010108)] // RPC_E_DISCONNECTED
    [InlineData(0x800401FD)] // CO_E_OBJNOTCONNECTED
    public void IsDisconnected_RecognisesDeadOutlookProxy(uint hresult)
    {
        Assert.True(OutlookSession.IsDisconnected(ComError(hresult)));
    }

    [Fact]
    public void IsDisconnected_FindsDisconnectWrappedAsFolderNotFound()
    {
        var wrapped = new OutlookMcpException(ErrorCodes.FolderNotFound, "Outlook could not find the requested folder.", "List folders again.",
            ComError(0x800706BA));
        Assert.True(OutlookSession.IsDisconnected(wrapped));
    }

    [Fact]
    public void IsDisconnected_IgnoresOrdinaryComFailures()
    {
        Assert.False(OutlookSession.IsDisconnected(ComError(0x8004010F)));
        Assert.False(OutlookSession.IsDisconnected(new InvalidOperationException()));
        Assert.False(OutlookSession.IsDisconnected(null));
    }

    private static Exception ComError(uint hresult)
    {
        var exception = Marshal.GetExceptionForHR(unchecked((int)hresult));
        Assert.IsType<COMException>(exception);
        return exception;
    }
}
