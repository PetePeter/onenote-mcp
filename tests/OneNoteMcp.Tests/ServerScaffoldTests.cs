using System.Reflection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol.Server;
using OneNoteMcp;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Scaffold-level protection: the host wires up cleanly and the advertised tool
/// set matches the v1 catalog. These are COM-free and run everywhere.
/// </summary>
public class ServerScaffoldTests
{
    [Fact]
    public void Host_builds_without_starting_transport()
    {
        // Build (not Run) so the stdio transport never starts; asserts DI wiring
        // for the MCP server + tool discovery resolves.
        using var host = ServerHost.CreateBuilder(Array.Empty<string>()).Build();
        Assert.NotNull(host.Services);
    }

    [Fact]
    public void Registers_exactly_the_expected_tool_set()
    {
        var toolNames = DiscoverToolNames(typeof(DiagnosticsTool).Assembly);

        Assert.Equal(
            new[]
            {
                "onenote_close_notebook",
                "onenote_create_notebook",
                "onenote_create_page",
                "onenote_create_section",
                "onenote_delete_node",
                "onenote_delete_page",
                "onenote_diagnostics",
                "onenote_export_one",
                "onenote_export_onepkg",
                "onenote_export_pdf",
                "onenote_extract_page_files",
                "onenote_find_pages",
                "onenote_get_hierarchy",
                "onenote_get_page",
                "onenote_get_page_info",
                "onenote_list_notebooks",
                "onenote_open_notebook",
                "onenote_rename_node",
                "onenote_update_page",
            },
            toolNames);
    }

    [Fact]
    public void Diagnostics_reports_server_version()
    {
        var result = DiagnosticsTool.OneNoteDiagnostics();

        Assert.StartsWith("OneNoteMcp server ", result);
        Assert.False(string.IsNullOrWhiteSpace(DiagnosticsTool.ServerVersion));
    }

    /// <summary>
    /// Enumerates every [McpServerTool] method under a [McpServerToolType] and
    /// returns its advertised name (explicit Name, else the method name).
    /// </summary>
    private static string[] DiscoverToolNames(Assembly assembly)
    {
        return assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance))
            .Select(m => (method: m, attr: m.GetCustomAttribute<McpServerToolAttribute>()))
            .Where(x => x.attr is not null)
            .Select(x => x.attr!.Name ?? x.method.Name)
            .OrderBy(n => n)
            .ToArray();
    }
}
