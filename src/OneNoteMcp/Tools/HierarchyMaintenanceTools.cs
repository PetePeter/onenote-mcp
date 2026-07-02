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
        [Description("OneNote object ID of the child node.")] string objectId) => ToolError.Guard(() =>
        OneNoteSession.Instance.GetHierarchyParent(objectId));

    [McpServerTool(Name = "onenote_get_special_location")]
    [Description("Returns the filesystem path of a OneNote special location: backup, unfiledNotes, or defaultNotebook.")]
    public static string GetSpecialLocation(
        [Description("Location: backup, unfiledNotes, or defaultNotebook.")] string location) => ToolError.Guard(() =>
        OneNoteSession.Instance.GetSpecialLocation(OneNoteEnumMapper.MapSpecialLocation(location)));

    [McpServerTool(Name = "onenote_find_meta")]
    [Description("Searches page metadata by name and returns hierarchy XML of matching pages.")]
    public static string FindMeta(
        [Description("Metadata name to search for.")] string searchName,
        [Description("Optional object ID to search under; empty for the root.")] string? startNodeId = null) => ToolError.Guard(() =>
        OneNoteSession.Instance.FindMeta(searchName, OneNoteXmlSchema.Xs2013, startNodeId ?? ""));

    [McpServerTool(Name = "onenote_merge_files")]
    [Description("Three-way merges OneNote files (base/client/server) into a target file.")]
    public static string MergeFiles(
        [Description("Path to the base file.")] string baseFile,
        [Description("Path to the client file.")] string clientFile,
        [Description("Path to the server file.")] string serverFile,
        [Description("Path to write the merged target file.")] string targetFile) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.MergeFiles(baseFile, clientFile, serverFile, targetFile);
        return "{\"merged\":true}";
    });

    [McpServerTool(Name = "onenote_merge_sections")]
    [Description("Merges the pages of a source section into a destination section.")]
    public static string MergeSections(
        [Description("OneNote object ID of the source section.")] string sourceId,
        [Description("OneNote object ID of the destination section.")] string destId) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.MergeSections(sourceId, destId);
        return "{\"merged\":true}";
    });

    [McpServerTool(Name = "onenote_sync_hierarchy")]
    [Description("Forces a sync of a hierarchy node (notebook/section).")]
    public static string SyncHierarchy(
        [Description("OneNote object ID of the node to sync.")] string hierarchyId) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.SyncHierarchy(hierarchyId);
        return "{\"synced\":true}";
    });

    [McpServerTool(Name = "onenote_set_filing_location")]
    [Description("Sets the section OneNote files a kind of Outlook item into. location: email|contacts|tasks|meetings|webContent|printOuts; locationType: namedSectionNewPage|currentSectionNewPage|currentPage|namedPage.")]
    public static string SetFilingLocation(
        [Description("Filing location: email, contacts, tasks, meetings, webContent, or printOuts.")] string location,
        [Description("Filing location type: namedSectionNewPage, currentSectionNewPage, currentPage, or namedPage.")] string locationType,
        [Description("OneNote object ID of the target section.")] string sectionId) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.SetFilingLocation(
            OneNoteEnumMapper.MapFilingLocation(location),
            OneNoteEnumMapper.MapFilingLocationType(locationType),
            sectionId);
        return "{\"set\":true}";
    });
}
