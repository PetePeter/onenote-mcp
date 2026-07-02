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
    public static string ListNotebooks(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version) =>
        ToolVersion.Guarded(version, Capability.GetHierarchy, s =>
            JsonSerializer.Serialize(HierarchyParser.ParseNotebooks(
                s.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013))));

    [McpServerTool(Name = "onenote_get_hierarchy")]
    [Description("Returns raw OneNote hierarchy XML for a node ID at the given scope (notebooks|sections|pages).")]
    public static string GetHierarchy(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the start node; empty string for the root.")] string nodeId,
        [Description("Scope: notebooks, sections, or pages.")] string scope) =>
        ToolVersion.Guarded(version, Capability.GetHierarchy, s =>
            s.GetHierarchy(nodeId ?? "", MapScope(scope), OneNoteXmlSchema.Xs2013));

    [McpServerTool(Name = "onenote_find_pages")]
    [Description("Searches OneNote for pages matching a query; returns matching pages with id, title, and section id.")]
    public static string FindPages(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Search query text.")] string query) =>
        ToolVersion.Guarded(version, Capability.FindPages, s =>
            JsonSerializer.Serialize(HierarchyParser.ParsePages(
                s.FindPages(query, OneNoteXmlSchema.Xs2013))));

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
