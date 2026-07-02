using System.Runtime.InteropServices;
using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>Unit tests for ComRetry — pure, no COM runtime required.</summary>
public sealed class ComRetryTests
{
    // ── IsTransient ───────────────────────────────────────────────────────────

    [Fact]
    public void IsTransient_RetryLater_ReturnsTrue()
    {
        Assert.True(ComRetry.IsTransient(unchecked((int)0x8001010A)));
    }

    [Fact]
    public void IsTransient_CallRejected_ReturnsTrue()
    {
        Assert.True(ComRetry.IsTransient(unchecked((int)0x80010001)));
    }

    [Fact]
    public void IsTransient_ServerUnavailable_ReturnsFalse()
    {
        // 0x800706BA means the COM object is dead and must be recreated — not retried
        Assert.False(ComRetry.IsTransient(unchecked((int)0x800706BA)));
    }

    [Fact]
    public void IsTransient_GenericError_ReturnsFalse()
    {
        Assert.False(ComRetry.IsTransient(unchecked((int)0x80004005)));
    }

    // ── Execute — success ─────────────────────────────────────────────────────

    [Fact]
    public void Execute_SucceedsFirstAttempt_ReturnsValue()
    {
        var result = ComRetry.Execute(() => 42, maxAttempts: 3, sleep: _ => { });
        Assert.Equal(42, result);
    }

    // ── Execute — retry on transient then succeed ─────────────────────────────

    [Fact]
    public void Execute_TransientThenSuccess_RetriesAndReturnsValue()
    {
        int callCount = 0;
        int sleepCount = 0;

        var result = ComRetry.Execute(() =>
        {
            callCount++;
            if (callCount < 3)
                throw new COMException("busy", unchecked((int)0x8001010A));
            return 99;
        }, maxAttempts: 5, sleep: _ => sleepCount++);

        Assert.Equal(99, result);
        Assert.Equal(3, callCount);
        Assert.Equal(2, sleepCount); // slept between attempt 1→2 and 2→3
    }

    // ── Execute — non-transient rethrows immediately ──────────────────────────

    [Fact]
    public void Execute_NonTransientCom_RethrowsImmediately()
    {
        int callCount = 0;

        var ex = Assert.Throws<COMException>(() =>
            ComRetry.Execute<int>(() =>
            {
                callCount++;
                throw new COMException("dead", unchecked((int)0x800706BA));
            }, maxAttempts: 5, sleep: _ => { })
        );

        Assert.Equal(unchecked((int)0x800706BA), ex.HResult);
        Assert.Equal(1, callCount); // must not retry non-transient
    }

    // ── Execute — gives up after maxAttempts ──────────────────────────────────

    [Fact]
    public void Execute_AllAttemptsTransient_ThrowsAfterMaxAttempts()
    {
        int callCount = 0;

        var ex = Assert.Throws<COMException>(() =>
            ComRetry.Execute<int>(() =>
            {
                callCount++;
                throw new COMException("busy", unchecked((int)0x8001010A));
            }, maxAttempts: 3, sleep: _ => { })
        );

        Assert.Equal(3, callCount);
        Assert.Equal(unchecked((int)0x8001010A), ex.HResult);
    }

    // ── Execute — non-COM exception rethrows immediately ─────────────────────

    [Fact]
    public void Execute_NonComException_RethrowsImmediately()
    {
        int callCount = 0;

        Assert.Throws<InvalidOperationException>(() =>
            ComRetry.Execute<int>(() =>
            {
                callCount++;
                throw new InvalidOperationException("oops");
            }, maxAttempts: 5, sleep: _ => { })
        );

        Assert.Equal(1, callCount);
    }

    // ── Void overload ─────────────────────────────────────────────────────────

    [Fact]
    public void Execute_VoidOverload_RetriesTransient()
    {
        int callCount = 0;

        ComRetry.Execute(() =>
        {
            callCount++;
            if (callCount < 2)
                throw new COMException("busy", unchecked((int)0x8001010A));
        }, maxAttempts: 3, sleep: _ => { });

        Assert.Equal(2, callCount);
    }
}
