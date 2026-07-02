using System.ComponentModel;
using System.Linq;
using System.Runtime.Versioning;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Discovery entry-point that lists every installed OneNote desktop version
/// with its CLSID, exe path, and supported capabilities. Takes no version
/// argument — it is the tool callers use to find version tokens.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class VersionTools
{
    [McpServerTool(Name = "onenote_list_versions")]
    [Description("Lists installed OneNote desktop versions with CLSID, exe path, and supported capabilities. Discovery entry point — takes no version argument.")]
    public static string ListVersions() =>
        ToolError.Guard(() => JsonSerializer.Serialize(BuildReport(OneNoteVersion.DetectAll())));

    /// <summary>
    /// Builds capability reports from detected install records. Internal so tests
    /// can inject a fake install list without touching the registry.
    /// </summary>
    internal static IReadOnlyList<VersionReport> BuildReport(IReadOnlyList<OneNoteInstall> installs) =>
        installs.Select(i => new VersionReport(
            i.DisplayName,
            i.Major,
            i.Clsid,
            i.ExePath,
            Default: false,
            OneNoteCapabilities.SupportedBy(i.Major)
                .Select(c => c.ToString())
                .OrderBy(x => x)
                .ToList()))
        .ToList();
}

/// <summary>
/// Report for a single installed OneNote version. Default is always false —
/// no default version exists; callers must pass an explicit version token.
/// </summary>
public sealed record VersionReport(
    string Version,
    int Major,
    string Clsid,
    string ExePath,
    bool Default,
    IReadOnlyList<string> Capabilities);
