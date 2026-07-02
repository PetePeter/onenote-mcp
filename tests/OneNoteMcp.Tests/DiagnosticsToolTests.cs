using System;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure, COM-free unit tests for the diagnostics report formatting. The live
/// <c>Collect()</c> path (process/COM probing) is exercised by the COM gate; here
/// we lock the pure <see cref="DiagnosticsTool.Format"/> shape so callers get a
/// complete, stable line regardless of machine state.
/// </summary>
public sealed class DiagnosticsToolTests
{
    [Fact]
    public void Format_IncludesAllFields()
    {
        var info = new DiagnosticsInfo(
            ServerVersion: "1.2.3",
            OneNoteVersion: "2016/2019/365",
            Running: true,
            OpenNotebookCount: 4,
            LastError: "OneNote error 0x80042014: The object does not exist.");

        var text = DiagnosticsTool.Format(info);

        Assert.StartsWith("OneNoteMcp server ", text);
        Assert.Contains("1.2.3", text);
        Assert.Contains("2016/2019/365", text);
        Assert.Contains("4", text);
        Assert.Contains("0x80042014", text);
    }

    [Fact]
    public void Format_NotRunning_ShowsNoRunningAndNaCount()
    {
        var info = new DiagnosticsInfo(
            ServerVersion: "1.0.0",
            OneNoteVersion: "not detected",
            Running: false,
            OpenNotebookCount: null,
            LastError: null);

        var text = DiagnosticsTool.Format(info);

        Assert.Contains("n/a", text);   // unknown open-notebook count when not running
        Assert.Contains("none", text);  // no last error
    }
}
