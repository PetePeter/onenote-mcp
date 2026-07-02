using System.Runtime.InteropServices;

namespace OneNoteMcp.Interop;

/// <summary>
/// Retry helper for transient COM errors.
/// OneNote's single-threaded COM apartment can temporarily reject calls while
/// processing another request; two well-known HRESULTs signal "try again soon".
/// </summary>
public static class ComRetry
{
    // RPC_E_CALL_REJECTED — COM server busy, can retry
    private const int RpcECallRejected = unchecked((int)0x80010001);
    // RPC_E_SERVERCALL_RETRYLATER — server says try later, can retry
    private const int RpcEServerCallRetryLater = unchecked((int)0x8001010A);

    /// <summary>
    /// Returns true if the HRESULT represents a transient COM busy condition
    /// that is safe to retry. Server-unavailable (0x800706BA) is NOT transient:
    /// the COM object must be recreated rather than retried.
    /// </summary>
    public static bool IsTransient(int hResult) =>
        hResult == RpcECallRejected || hResult == RpcEServerCallRetryLater;

    /// <summary>
    /// Executes <paramref name="action"/> retrying on transient COM errors.
    /// Non-transient <see cref="COMException"/> and all other exceptions are
    /// rethrown immediately. After <paramref name="maxAttempts"/> transient
    /// failures the last exception is rethrown.
    /// </summary>
    /// <param name="action">The COM-touching work to execute.</param>
    /// <param name="maxAttempts">Maximum total attempts (default 5).</param>
    /// <param name="sleep">
    /// Called with the 1-based attempt number before each retry.
    /// Defaults to <c>Thread.Sleep(200 * attempt)</c>. Inject a no-op in tests.
    /// </param>
    public static T Execute<T>(Func<T> action, int maxAttempts = 5, Action<int>? sleep = null)
    {
        sleep ??= attempt => Thread.Sleep(200 * attempt);

        COMException? last = null;
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                return action();
            }
            catch (COMException ex) when (IsTransient(ex.HResult))
            {
                last = ex;
                if (attempt < maxAttempts)
                    sleep(attempt);
            }
            // Non-transient COMException or any other exception: propagate immediately
        }

        throw last!;
    }

    /// <summary>
    /// Executes <paramref name="action"/> retrying on transient COM errors.
    /// Delegates to the generic overload.
    /// </summary>
    public static void Execute(Action action, int maxAttempts = 5, Action<int>? sleep = null) =>
        Execute<int>(() => { action(); return 0; }, maxAttempts, sleep);
}
