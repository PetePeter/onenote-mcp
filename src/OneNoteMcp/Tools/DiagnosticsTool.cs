using System.ComponentModel;
using System.Reflection;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;

namespace OneNoteMcp.Tools;

/// <summary>
/// Diagnostics tool reporting server health, detected OneNote version, and COM
/// availability. Does NOT launch OneNote — version detection is registry-only.
/// </summary>
[McpServerToolType]
public static class DiagnosticsTool
{
    [McpServerTool(Name = "onenote_diagnostics")]
    [Description("Returns OneNote MCP server diagnostics including server version, detected OneNote version, and COM availability.")]
    public static string OneNoteDiagnostics()
    {
        var oneNoteVersion = OneNoteVersion.Detect()?.DisplayName ?? "not detected";
        var comAvailable = OneNoteSession.IsComAvailable ? "available" : "unavailable";
        return $"OneNoteMcp server {ServerVersion} | OneNote {oneNoteVersion} | COM {comAvailable}";
    }

    /// <summary>Informational assembly version, e.g. "1.0.0".</summary>
    public static string ServerVersion =>
        typeof(DiagnosticsTool).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(DiagnosticsTool).Assembly.GetName().Version?.ToString()
        ?? "0.0.0";
}
