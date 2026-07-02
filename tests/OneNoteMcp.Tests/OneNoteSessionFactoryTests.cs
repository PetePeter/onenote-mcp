using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fakes;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Unit tests for OneNoteSession.For(clsid) factory: CLSID-keyed caching and
/// distinct-instance behaviour. Uses AppFactoryOverride so no real COM is needed.
/// </summary>
[Collection("OneNote COM")]
public sealed class OneNoteSessionFactoryTests : IDisposable
{
    // The canonical CLSID for OneNote 2016 (also used for 2013/2019/365 in the catalog).
    private const string ClsidA = "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}";
    // The canonical CLSID for OneNote 2010.
    private const string ClsidB = "{D7FAC39E-7FF1-49AA-98CF-A1DDD316337E}";

    public OneNoteSessionFactoryTests() => OneNoteSession.ResetForTests();
    public void Dispose() => OneNoteSession.ResetForTests();

    [Fact]
    public void For_SameClsid_ReturnsCachedInstance()
    {
        OneNoteSession.AppFactoryOverride = _ => new FakeOneNoteApp();

        var s1 = OneNoteSession.For(ClsidA);
        var s2 = OneNoteSession.For(ClsidA);

        Assert.Same(s1, s2);
    }

    [Fact]
    public void For_DifferentClsid_ReturnsDistinctInstances()
    {
        OneNoteSession.AppFactoryOverride = _ => new FakeOneNoteApp();

        var sA = OneNoteSession.For(ClsidA);
        var sB = OneNoteSession.For(ClsidB);

        Assert.NotSame(sA, sB);
    }
}
