using System.Runtime.InteropServices;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fakes;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM lifetime and apartment ownership of <see cref="OneNoteSession"/> (issue #1):
/// the owned app is released exactly once, on the same STA thread that activated
/// and called it, cached sessions are disposed on host stop, and no work runs
/// after release.
/// </summary>
[Collection("OneNote COM")]
public sealed class OneNoteSessionLifetimeTests : IDisposable
{
    public OneNoteSessionLifetimeTests() => OneNoteSession.ResetForTests();
    public void Dispose() => OneNoteSession.ResetForTests();

    private static string Read(OneNoteSession s) =>
        s.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013);

    [Fact]
    public void Dispose_ReleasesApp_ExactlyOnce()
    {
        var app = new FakeOneNoteApp();
        var session = new OneNoteSession(() => app);
        Read(session);

        session.Dispose();
        session.Dispose();

        Assert.Equal(1, app.DisposeCount);
    }

    [Fact]
    public void Dispose_UnusedSession_NeverActivatesCom()
    {
        int factoryCalls = 0;
        var session = new OneNoteSession(() => { factoryCalls++; return new FakeOneNoteApp(); });

        session.Dispose();

        Assert.Equal(0, factoryCalls);
    }

    [Fact]
    public void ActivationCallsAndRelease_AllRunOnOneStaThread()
    {
        var app = new FakeOneNoteApp();
        int? factoryThread = null;
        ApartmentState? factoryApartment = null;
        var session = new OneNoteSession(() =>
        {
            factoryThread = Environment.CurrentManagedThreadId;
            factoryApartment = Thread.CurrentThread.GetApartmentState();
            return app;
        });

        // Calls from several thread-pool threads, as the MCP SDK would make them.
        Task.WaitAll(Enumerable.Range(0, 4).Select(_ => Task.Run(() => Read(session))).ToArray());
        session.Dispose();

        Assert.Equal(ApartmentState.STA, factoryApartment);
        Assert.NotEqual(Environment.CurrentManagedThreadId, factoryThread);
        Assert.All(app.CallThreadIds, id => Assert.Equal(factoryThread, id));
        Assert.Equal(factoryThread, app.DisposeThreadId);
    }

    [Fact]
    public void CallAfterDispose_ThrowsObjectDisposed()
    {
        var session = new OneNoteSession(() => new FakeOneNoteApp());
        session.Dispose();

        Assert.Throws<ObjectDisposedException>(() => Read(session));
    }

    [Fact]
    public void Dispose_DrainsInFlightCall_BeforeRelease()
    {
        using var entered = new ManualResetEventSlim();
        using var gate = new ManualResetEventSlim();
        var app = new FakeOneNoteApp { OnCall = () => { entered.Set(); gate.Wait(); } };
        var session = new OneNoteSession(() => app);

        var call = Task.Run(() => Read(session));
        Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));

        var dispose = Task.Run(session.Dispose);
        Assert.False(dispose.Wait(200)); // must wait for the in-flight call
        Assert.Equal(0, app.DisposeCount);

        gate.Set();
        Assert.True(dispose.Wait(TimeSpan.FromSeconds(5)));
        call.GetAwaiter().GetResult(); // in-flight call completed normally
        Assert.Equal(1, app.DisposeCount);
        Assert.Single(app.CallThreadIds); // nothing ran after release
    }

    [Fact]
    public void DeadProxy_ReleasesOldApp_BeforeRecreating()
    {
        var dead = new FakeOneNoteApp { DeadProxy = true };
        var live = new FakeOneNoteApp { ReturnValue = "<Notebooks />" };
        var apps = new Queue<FakeOneNoteApp>(new[] { dead, live });
        using var session = new OneNoteSession(() => apps.Dequeue());

        Assert.Throws<COMException>(() => Read(session));
        Assert.Equal(1, dead.DisposeCount);

        Assert.Equal("<Notebooks />", Read(session));
        Assert.Equal(0, live.DisposeCount);
    }

    [Fact]
    public void Activation_RetriesWithBackoff_WhileOneNoteIsShuttingDown()
    {
        var waits = new List<TimeSpan>();
        OneNoteSession.ActivationBackoff = waits.Add;
        int attempts = 0;
        using var session = new OneNoteSession(() =>
            ++attempts < 3
                ? throw new COMException("exec failure", unchecked((int)0x80080005))
                : new FakeOneNoteApp { ReturnValue = "<Notebooks />" });

        Assert.Equal("<Notebooks />", Read(session));
        Assert.Equal(new[] { TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(2) }, waits);
    }

    [Fact]
    public void Activation_GivesUp_AfterFourAttempts()
    {
        var waits = new List<TimeSpan>();
        OneNoteSession.ActivationBackoff = waits.Add;
        int attempts = 0;
        using var session = new OneNoteSession(() =>
        {
            attempts++;
            throw new COMException("exec failure", unchecked((int)0x80080005));
        });

        var ex = Assert.Throws<COMException>(() => Read(session));
        Assert.Equal(unchecked((int)0x80080005), ex.HResult);
        Assert.Equal(4, attempts);
        Assert.Equal(3, waits.Count);
    }

    [Fact]
    public async Task HostStop_DisposesCachedSessions()
    {
        var app = new FakeOneNoteApp();
        OneNoteSession.AppFactoryOverride = _ => app;
        var session = OneNoteSession.For(OneNoteVersionCatalog.All[0].Clsid);
        Read(session);

        using var host = ServerHost.CreateBuilder(Array.Empty<string>()).Build();
        foreach (var hosted in host.Services.GetServices<IHostedService>().OfType<OneNoteSessionLifetime>())
            await hosted.StopAsync(CancellationToken.None);

        Assert.Equal(1, app.DisposeCount);
        Assert.Throws<ObjectDisposedException>(() => Read(session));
        Assert.NotSame(session, OneNoteSession.For(OneNoteVersionCatalog.All[0].Clsid));
    }
}
