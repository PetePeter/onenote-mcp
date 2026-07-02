using OneNoteMcp.Interop;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure, COM-free tests for VersionTools.BuildReport. No COM override needed
/// because the method takes an injected install list (no registry access).
/// </summary>
public sealed class VersionToolsTests
{
    [Fact]
    public void BuildReport_TwoInstalls_MapsCapabilities()
    {
        var installs = new[]
        {
            new OneNoteInstall(12, "2007",  "{0039FFEC-A022-4232-8274-6B34787BFC27}", @"C:\Program Files\Office12\ONENOTE.EXE"),
            new OneNoteInstall(16, "2016",  "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}", @"C:\Program Files\Office16\ONENOTE.EXE"),
        };

        var reports = VersionTools.BuildReport(installs);

        Assert.Equal(2, reports.Count);

        var v12 = Assert.Single(reports, r => r.Major == 12);
        Assert.Equal(17, v12.Capabilities.Count);
        Assert.DoesNotContain("MergeFiles", v12.Capabilities);
        Assert.DoesNotContain("SyncHierarchy", v12.Capabilities);
        Assert.False(v12.Default);

        var v16 = Assert.Single(reports, r => r.Major == 16);
        Assert.Equal(23, v16.Capabilities.Count);
        Assert.Contains("SyncHierarchy", v16.Capabilities);
        Assert.False(v16.Default);
    }
}
