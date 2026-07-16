using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.OutlookInterop;

/// <summary>Serializes all Outlook COM access onto one dedicated STA thread.</summary>
public sealed class OutlookStaDispatcher : IAsyncDisposable
{
    private readonly Channel<IWorkItem> _queue = Channel.CreateUnbounded<IWorkItem>(new UnboundedChannelOptions { SingleReader = true, SingleWriter = false });
    private readonly CancellationTokenSource _shutdown = new();
    private readonly Thread _thread;
    private readonly TimeSpan _timeout;
    private readonly ILogger<OutlookStaDispatcher> _logger;

    public OutlookStaDispatcher(OutlookMcpOptions options, ILogger<OutlookStaDispatcher> logger)
    {
        _logger = logger;
        _timeout = TimeSpan.FromSeconds(Math.Clamp(options.Outlook.OperationTimeoutSeconds, 1, 300));
        _thread = new Thread(Run) { IsBackground = true, Name = "Outlook MCP STA" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public async Task<T> InvokeAsync<T>(Func<T> operation, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_shutdown.IsCancellationRequested, this);
        using var timeout = new CancellationTokenSource(_timeout);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeout.Token, _shutdown.Token);
        var item = new WorkItem<T>(operation);
        await _queue.Writer.WriteAsync(item, linked.Token).ConfigureAwait(false);
        try
        {
            return await item.Task.WaitAsync(linked.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException ex) when (timeout.IsCancellationRequested)
        {
            throw new OutlookMcpException(ErrorCodes.SearchTimeout, "The Outlook operation exceeded its configured timeout.", "Narrow the request scope or retry after Outlook finishes synchronising.", ex);
        }
        catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested)
        {
            throw new OutlookMcpException(ErrorCodes.OperationCancelled, "The Outlook operation was cancelled.", "Retry the operation if it is still needed.", ex);
        }
    }

    private void Run()
    {
        try
        {
            while (_queue.Reader.WaitToReadAsync(_shutdown.Token).AsTask().GetAwaiter().GetResult())
            {
                while (_queue.Reader.TryRead(out var item)) item.Execute();
            }
        }
        catch (OperationCanceledException) when (_shutdown.IsCancellationRequested) { }
        catch (Exception ex) { _logger.LogCritical(ex, "The Outlook STA dispatcher terminated unexpectedly"); }
    }

    public ValueTask DisposeAsync()
    {
        if (_shutdown.IsCancellationRequested) return ValueTask.CompletedTask;
        _queue.Writer.TryComplete();
        _shutdown.Cancel();
        if (!_thread.Join(TimeSpan.FromSeconds(5))) _logger.LogWarning("Outlook STA dispatcher did not stop within five seconds");
        _shutdown.Dispose();
        return ValueTask.CompletedTask;
    }

    private interface IWorkItem { void Execute(); }

    private sealed class WorkItem<T>(Func<T> operation) : IWorkItem
    {
        private readonly TaskCompletionSource<T> _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<T> Task => _completion.Task;
        public void Execute()
        {
            try { _completion.TrySetResult(operation()); }
            catch (Exception ex) { _completion.TrySetException(ex); }
        }
    }
}
