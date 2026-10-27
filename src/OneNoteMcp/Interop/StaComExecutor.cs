using System.Collections.Concurrent;
using System.Runtime.Versioning;

namespace OneNoteMcp.Interop;

/// <summary>
/// Runs work items one at a time, in order, on a single dedicated STA thread.
/// Gives a COM object one owner: activation, every call and the final release
/// all happen on the same apartment thread (issue #1). The thread starts lazily,
/// so an unused executor costs nothing.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class StaComExecutor : IDisposable
{
    private readonly BlockingCollection<Action> _queue = new();
    private readonly object _startLock = new();
    private Thread? _thread;

    /// <summary>
    /// Runs <paramref name="work"/> on the STA thread and blocks until it finishes,
    /// rethrowing any exception it throws.
    /// Throws <see cref="ObjectDisposedException"/> once the executor is disposed.
    /// </summary>
    public T Invoke<T>(Func<T> work)
    {
        if (Thread.CurrentThread == _thread)
            return work(); // Re-entrant call from the owner thread: run inline.

        EnsureStarted();
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        try
        {
            _queue.Add(() =>
            {
                try { tcs.SetResult(work()); }
                catch (Exception ex) { tcs.SetException(ex); }
            });
        }
        catch (Exception ex) when (ex is InvalidOperationException or ObjectDisposedException)
        {
            throw new ObjectDisposedException(nameof(StaComExecutor));
        }

        return tcs.Task.GetAwaiter().GetResult(); // rethrows the original exception
    }

    private void EnsureStarted()
    {
        if (_thread is not null) return;
        lock (_startLock)
        {
            if (_thread is not null || _queue.IsAddingCompleted) return;
            var thread = new Thread(Run) { IsBackground = true, Name = "OneNote COM (STA)" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            _thread = thread;
        }
    }

    private void Run()
    {
        foreach (var work in _queue.GetConsumingEnumerable())
            work();
    }

    /// <summary>
    /// Stops accepting work, drains queued items, then ends the thread. Idempotent.
    /// Must not be called from the owner thread.
    /// </summary>
    public void Dispose()
    {
        lock (_startLock)
        {
            if (_queue.IsAddingCompleted) return;
            _queue.CompleteAdding();
        }
        _thread?.Join();
    }
}
