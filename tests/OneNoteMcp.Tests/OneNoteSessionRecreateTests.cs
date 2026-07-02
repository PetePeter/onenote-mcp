using System.Runtime.InteropServices;
using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fakes;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Unit tests for OneNoteSession's recreate-on-death behaviour.
/// Uses injected IOneNoteApp fakes — no real OneNote required.
/// </summary>
[Collection("OneNote COM")]
public sealed class OneNoteSessionRecreateTests : IDisposable
{
    public OneNoteSessionRecreateTests() => OneNoteSession.ResetForTests();
    public void Dispose() => OneNoteSession.ResetForTests();

    [Fact]
    public void FactoryCalledAgain_AfterDeadProxy()
    {
        int factoryCallCount = 0;

        // First factory call returns a dead proxy (throws 0x800706BA).
        // Second factory call returns a live app (succeeds).
        using var session = new OneNoteSession(() =>
        {
            factoryCallCount++;
            return factoryCallCount == 1
                ? new FakeOneNoteApp { DeadProxy = true }
                : new FakeOneNoteApp { ReturnValue = "<Notebooks />" };
        });

        Assert.Equal(0, factoryCallCount); // lazy: nothing created yet

        // First call: factory builds dead app (#1), which throws 0x800706BA.
        // Session marks proxy dead (invalidates). COMException propagates out.
        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        Assert.Equal(1, factoryCallCount); // factory was called once for dead app

        // Second call: factory builds live app (#2), which succeeds.
        var xml = session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013);

        Assert.Equal(2, factoryCallCount); // factory called again — recreate-on-death confirmed
        Assert.Equal("<Notebooks />", xml);
    }

    [Fact]
    public void InvalidateForTests_CausesRecreationOnNextCall()
    {
        int factoryCallCount = 0;

        using var session = new OneNoteSession(() =>
        {
            factoryCallCount++;
            // Always return a dead app so we can observe recreation attempts
            return new FakeOneNoteApp { DeadProxy = true };
        });

        // Force invalidation without making any COM call
        session.InvalidateForTests();

        // Next call: factory creates dead app, which throws 0x800706BA
        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        // Factory was called exactly once to create the (dead) app
        Assert.Equal(1, factoryCallCount);

        // After the failure the app is invalidated again — next call will recreate
        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        Assert.Equal(2, factoryCallCount); // confirmed: each dead-proxy failure triggers recreation
    }

    [Fact]
    public void LastComError_IsRecorded_AfterComFailure()
    {
        using var session = new OneNoteSession(() => new FakeOneNoteApp { DeadProxy = true });

        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        // The session records the most recent COM failure (mapped) so diagnostics
        // can surface a "last error" without any logging infrastructure.
        Assert.False(string.IsNullOrWhiteSpace(OneNoteSession.LastComError));
        Assert.Contains("0x800706BA", OneNoteSession.LastComError!, StringComparison.OrdinalIgnoreCase);
    }
}
