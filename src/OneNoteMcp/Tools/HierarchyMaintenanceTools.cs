using System.ComponentModel;
using System.Runtime.Versioning;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Maintenance and query tools over the OneNote hierarchy: parent lookup, special
/// locations, metadata search, merges, sync, and the Outlook filing location.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class HierarchyMaintenanceTools
{
    [McpServerTool(Name = "onenote_get_hierarchy_parent")]
    [Description("Returns the object ID of the parent of a hierarchy node.")]
    public static string GetHierarchyParent(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the child node.")] string objectId) =>
        ToolVersion.Guarded(version, Capability.GetHierarchyParent, s =>
            s.GetHierarchyParent(objectId));

    [McpServerTool(Name = "onenote_get_special_location")]
    [Description("Returns the filesystem path of a OneNote special location: backup, unfiledNotes, or defaultNotebook.")]
    public static string GetSpecialLocation(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Location: backup, unfiledNotes, or defaultNotebook.")] string location) =>
        ToolVersion.Guarded(version, Capability.GetSpecialLocation, s =>
            s.GetSpecialLocation(OneNoteEnumMapper.MapSpecialLocation(location)));

    [McpServerTool(Name = "onenote_find_meta")]
    [Description("Searches page metadata by name and returns hierarchy XML of matching pages.")]
    public static string FindMeta(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Metadata name to search for.")] string searchName,
        [Description("Optional object ID to search under; empty for the root.")] string? startNodeId = null) =>
        ToolVersion.Guarded(version, Capability.FindMeta, s =>
            s.FindMeta(searchName, OneNoteXmlSchema.Xs2013, startNodeId ?? ""));

    [McpServerTool(Name = "onenote_merge_files")]
    [Description("Three-way merges OneNote files (base/client/server) into a target file.")]
    public static string MergeFiles(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Path to the base file.")] string baseFile,
        [Description("Path to the client file.")] string clientFile,
        [Description("Path to the server file.")] string serverFile,
        [Description("Path to write the merged target file.")] string targetFile) =>
        ToolVersion.Guarded(version, Capability.MergeFiles, s =>
        {
            s.MergeFiles(baseFile, clientFile, serverFile, targetFile);
            return "{\"merged\":true}";
        });

    [McpServerTool(Name = "onenote_merge_sections")]
    [Description("Merges the pages of a source section into a destination section.")]
    public static string MergeSections(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the source section.")] string sourceId,
        [Description("OneNote object ID of the destination section.")] string destId) =>
        ToolVersion.Guarded(version, Capability.MergeSections, s =>
        {
            s.MergeSections(sourceId, destId);
            return "{\"merged\":true}";
        });

    [McpServerTool(Name = "onenote_sync_hierarchy")]
    [Description("Forces a sync of a hierarchy node (notebook/section).")]
    public static string SyncHierarchy(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the node to sync.")] string hierarchyId) =>
        ToolVersion.Guarded(version, Capability.SyncHierarchy, s =>
        {
            s.SyncHierarchy(hierarchyId);
            return "{\"synced\":true}";
        });

    [McpServerTool(Name = "onenote_set_filing_location")]
    [Description("Sets the section OneNote files a kind of Outlook item into. location: email|contacts|tasks|meetings|webContent|printOuts; locationType: namedSectionNewPage|currentSectionNewPage|currentPage|namedPage.")]
    public static string SetFilingLocation(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Filing location: email, contacts, tasks, meetings, webContent, or printOuts.")] string location,
        [Description("Filing location type: namedSectionNewPage, currentSectionNewPage, currentPage, or namedPage.")] string locationType,
        [Description("OneNote object ID of the target section.")] string sectionId) =>
        ToolVersion.Guarded(version, Capability.SetFilingLocation, s =>
        {
            s.SetFilingLocation(
                OneNoteEnumMapper.MapFilingLocation(location),
                OneNoteEnumMapper.MapFilingLocationType(locationType),
                sectionId);
            return "{\"set\":true}";
        });
}
