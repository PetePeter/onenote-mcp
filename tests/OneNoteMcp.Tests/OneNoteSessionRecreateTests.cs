using System.Runtime.InteropServices;
using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Unit tests for OneNoteSession's recreate-on-death behaviour.
/// Uses injected fake COM objects — no real OneNote required.
/// </summary>
[Collection("OneNote COM")]
public sealed class OneNoteSessionRecreateTests
{
    /// <summary>
    /// Fake that always throws 0x800706BA (server unavailable) on any method call.
    /// Used to simulate a dead COM proxy.
    /// </summary>
    private sealed class DeadApp
    {
        public void GetHierarchy(string startNodeId, int scope, ref string xml, int xmlSchema)
            => throw new COMException("Server unavailable", unchecked((int)0x800706BA));
    }

    /// <summary>
    /// Fake that always succeeds, returning a valid XML hierarchy string.
    /// Used as the replacement after a dead proxy is recreated.
    /// </summary>
    private sealed class LiveApp
    {
        public void GetHierarchy(string startNodeId, int scope, ref string xml, int xmlSchema)
            => xml = "<Notebooks />";
    }

    [Fact]
    public void FactoryCalledAgain_AfterDeadProxy()
    {
        int factoryCallCount = 0;

        // First factory call returns DeadApp (throws 0x800706BA).
        // Second factory call returns LiveApp (succeeds).
        using var session = new OneNoteSession(() =>
        {
            factoryCallCount++;
            return factoryCallCount == 1 ? (object)new DeadApp() : new LiveApp();
        });

        Assert.Equal(0, factoryCallCount); // lazy: nothing created yet

        // First call: factory builds DeadApp (#1), which throws 0x800706BA.
        // Session marks proxy dead (invalidates). COMException propagates out.
        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        Assert.Equal(1, factoryCallCount); // factory was called once for DeadApp

        // Second call: factory builds LiveApp (#2), which succeeds.
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
            return new DeadApp();
        });

        // Force invalidation without making any COM call
        session.InvalidateForTests();

        // Next call: factory creates DeadApp, which throws 0x800706BA
        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        // Factory was called exactly once to create the (dead) app
        Assert.Equal(1, factoryCallCount);

        // After the failure the app is invalidated again — next call will recreate
        Assert.Throws<COMException>(() =>
            session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013));

        Assert.Equal(2, factoryCallCount); // confirmed: each dead-proxy failure triggers recreation
    }
}
