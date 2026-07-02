using System.ComponentModel;
using System.Runtime.Versioning;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Read-only tools that surface the OneNote hierarchy: listing notebooks,
/// returning raw hierarchy XML for a node, and searching for pages.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class HierarchyTools
{
    [McpServerTool(Name = "onenote_list_notebooks")]
    [Description("Lists all open OneNote notebooks with their object IDs, names, paths, and flags.")]
    public static string ListNotebooks()
    {
        var xml = OneNoteSession.Instance.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013);
        return JsonSerializer.Serialize(HierarchyParser.ParseNotebooks(xml));
    }

    [McpServerTool(Name = "onenote_get_hierarchy")]
    [Description("Returns raw OneNote hierarchy XML for a node ID at the given scope (notebooks|sections|pages).")]
    public static string GetHierarchy(
        [Description("OneNote object ID of the start node; empty string for the root.")] string nodeId,
        [Description("Scope: notebooks, sections, or pages.")] string scope)
    {
        var mappedScope = MapScope(scope);
        return OneNoteSession.Instance.GetHierarchy(nodeId ?? "", mappedScope, OneNoteXmlSchema.Xs2013);
    }

    [McpServerTool(Name = "onenote_find_pages")]
    [Description("Searches OneNote for pages matching a query; returns matching pages with id, title, and section id.")]
    public static string FindPages(
        [Description("Search query text.")] string query)
    {
        var xml = OneNoteSession.Instance.FindPages(query, OneNoteXmlSchema.Xs2013);
        return JsonSerializer.Serialize(HierarchyParser.ParsePages(xml));
    }

    /// <summary>Maps a caller-supplied scope name to its OneNote HierarchyScope value.</summary>
    private static int MapScope(string scope) => scope?.ToLowerInvariant() switch
    {
        "notebooks" => OneNoteScope.HsNotebooks,
        "sections"  => OneNoteScope.HsSections,
        "pages"     => OneNoteScope.HsPages,
        "self"      => OneNoteScope.HsSelf,
        "children"  => OneNoteScope.HsChildren,
        _ => throw new ArgumentException(
            $"Unknown scope '{scope}'. Expected one of: notebooks, sections, pages, self, children.", nameof(scope)),
    };
}
