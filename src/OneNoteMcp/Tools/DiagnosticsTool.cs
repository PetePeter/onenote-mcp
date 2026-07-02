using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.Versioning;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>Immutable snapshot of server/OneNote health for a single diagnostics report.</summary>
public sealed record DiagnosticsInfo(
    string ServerVersion,
    string OneNoteVersion,
    bool Running,
    int? OpenNotebookCount,
    string? LastError);

/// <summary>
/// Diagnostics tool reporting server health, detected OneNote version, and COM
/// availability. Never launches OneNote — version detection is registry-only and
/// the COM probe runs only when OneNote is already running.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class DiagnosticsTool
{
    [McpServerTool(Name = "onenote_diagnostics")]
    [Description("Returns OneNote MCP server diagnostics including server version, detected OneNote version, running state, open notebook count, and last error.")]
    public static string OneNoteDiagnostics(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version) =>
        Format(Collect(version));

    /// <summary>
    /// Renders a diagnostics snapshot to a single stable line. Pure: no process or
    /// COM access, so callers get a complete report regardless of machine state.
    /// </summary>
    public static string Format(DiagnosticsInfo info) =>
        $"OneNoteMcp server {info.ServerVersion} | OneNote {info.OneNoteVersion} | running {(info.Running ? "yes" : "no")} | open notebooks {(info.OpenNotebookCount?.ToString() ?? "n/a")} | last error {info.LastError ?? "none"}";

    /// <summary>
    /// Gathers live diagnostics. The open-notebook probe runs only when OneNote is
    /// already running and COM is available, so it never launches OneNote or trips
    /// the first-run modal dialog.
    /// </summary>
    private static DiagnosticsInfo Collect(string version)
    {
        var oneNoteVersion = OneNoteVersion.Detect()?.DisplayName ?? "not detected";
        var running = System.Diagnostics.Process.GetProcessesByName("ONENOTE").Length > 0;

        int? openNotebookCount = null;
        string? probeError = null;

        // Defensively resolve the version — a bad token must not crash diagnostics.
        string? clsid = null;
        try { (clsid, _) = ToolVersion.Resolve(version); } catch { /* leave clsid null */ }

        if (running && OneNoteSession.IsComAvailable && clsid is not null)
        {
            try
            {
                var xml = OneNoteSession.For(clsid).GetHierarchy(
                    "", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013);
                openNotebookCount = HierarchyParser.ParseNotebooks(xml).Count;
            }
            catch (Exception ex)
            {
                probeError = ComErrorMapper.Describe(ex);
            }
        }

        return new DiagnosticsInfo(
            ServerVersion,
            oneNoteVersion,
            running,
            openNotebookCount,
            probeError ?? OneNoteSession.LastComError);
    }

    /// <summary>Informational assembly version, e.g. "1.0.0".</summary>
    public static string ServerVersion =>
        typeof(DiagnosticsTool).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(DiagnosticsTool).Assembly.GetName().Version?.ToString()
        ?? "0.0.0";
}
