using System.ComponentModel;
using System.Reflection;
using ModelContextProtocol.Server;

namespace OneNoteMcp.Tools;

/// <summary>
/// Placeholder diagnostics tool. The v1 catalog expands this to report the
/// detected OneNote version, running state, open notebook count and last error
/// (see plan P-0537). For the scaffold it reports only the server version.
/// </summary>
[McpServerToolType]
public static class DiagnosticsTool
{
    [McpServerTool(Name = "onenote_diagnostics")]
    [Description("Returns OneNote MCP server diagnostics. In the scaffold this is the server version only.")]
    public static string OneNoteDiagnostics() => $"OneNoteMcp server {ServerVersion}";

    /// <summary>Informational assembly version, e.g. "1.0.0".</summary>
    public static string ServerVersion =>
        typeof(DiagnosticsTool).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(DiagnosticsTool).Assembly.GetName().Version?.ToString()
        ?? "0.0.0";
}
