using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for <see cref="OneNoteVersion.DetectAll"/>. The registry
/// LocalServer32 lookup is injected as a fake so multiple installs — including the
/// 2010/2013 compat CLSIDs that both resolve to the modern exe — can be exercised
/// without touching the real registry or COM. Dedupe collapses CLSIDs whose exe
/// path is identical, keeping the most modern major.
/// </summary>
public sealed class OneNoteVersionDetectAllTests
{
    const string Clsid2007 = "{0039FFEC-A022-4232-8274-6B34787BFC27}";
    const string Clsid2010 = "{D7FAC39E-7FF1-49AA-98CF-A1DDD316337E}";
    const string ClsidModern = "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}";
    const string Office12 = @"C:\Program Files (x86)\Microsoft Office\Office12\ONENOTE.EXE";
    const string Office16 = @"C:\Program Files (x86)\Microsoft Office\Root\Office16\ONENOTE.EXE";

    static Func<string, string?> Registry(Dictionary<string, string> rows)
        => clsid => rows.TryGetValue(clsid, out var path) ? path : null;

    [Fact]
    public void DetectAll_MultipleInstalls_DedupesSameExe()
    {
        var rows = new Dictionary<string, string>
        {
            [Clsid2007] = Office12,   // 2007 → its own exe
            [Clsid2010] = Office16,   // 2010 compat stub → modern exe
            [ClsidModern] = Office16, // 2016 desktop → same modern exe
        };

        var installs = OneNoteVersion.DetectAll(Registry(rows));

        Assert.Equal(2, installs.Count);

        var v12 = Assert.Single(installs, i => i.Major == 12);
        Assert.Equal(Office12, v12.ExePath);
        Assert.Equal(Clsid2007, v12.Clsid);
        Assert.Equal("2007", v12.DisplayName);

        var v16 = Assert.Single(installs, i => i.Major == 16);
        Assert.Equal(Office16, v16.ExePath);
        Assert.Equal(ClsidModern, v16.Clsid); // most-modern CLSID survives the collapse
        Assert.Contains("2016", v16.DisplayName);
    }

    [Fact]
    public void DetectAll_NoInstalls_ReturnsEmpty()
        => Assert.Empty(OneNoteVersion.DetectAll(_ => null));

    [Fact]
    public void DetectAll_SingleInstall_ReturnsOneEntry()
    {
        var rows = new Dictionary<string, string> { [Clsid2007] = Office12 };

        var only = Assert.Single(OneNoteVersion.DetectAll(Registry(rows)));

        Assert.Equal(12, only.Major);
        Assert.Equal(Office12, only.ExePath);
    }

    [Fact]
    public void DetectAll_CaseInsensitiveExePaths_StillDeduped()
    {
        var rows = new Dictionary<string, string>
        {
            [Clsid2010] = Office16.ToLowerInvariant(),
            [ClsidModern] = Office16.ToUpperInvariant(),
        };

        var only = Assert.Single(OneNoteVersion.DetectAll(Registry(rows)));

        Assert.Equal(16, only.Major);
    }
}
