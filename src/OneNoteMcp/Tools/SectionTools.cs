using System;
using System.ComponentModel;
using System.Runtime.Versioning;
using System.Xml.Linq;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;

namespace OneNoteMcp.Tools;

/// <summary>
/// Write tools that create, rename, and delete OneNote hierarchy nodes
/// (sections, section groups, notebooks). Per the project's "no guardrails on
/// destructive ops" decision these apply without a conflict check.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class SectionTools
{
    [McpServerTool(Name = "onenote_create_section")]
    [Description("Creates a new section under a notebook or section group. Returns the new section's object ID.")]
    public static string CreateSection(
        [Description("OneNote object ID of the parent notebook or section group.")] string parentId,
        [Description("Name of the new section (a .one extension is added if absent).")] string name)
    {
        var sectionPath = name.EndsWith(".one", StringComparison.OrdinalIgnoreCase)
            ? name
            : name + ".one";

        return OneNoteSession.Instance.OpenHierarchy(
            sectionPath, parentId, OneNoteCreateFileType.CftSection);
    }

    [McpServerTool(Name = "onenote_rename_node")]
    [Description("Renames a section, section group, or notebook by its object ID.")]
    public static string RenameNode(
        [Description("OneNote object ID of the node to rename.")] string id,
        [Description("New name for the node.")] string newName)
    {
        // Fetch the node's own XML so the change fragment carries the correct
        // schema namespace; mutate the parsed element in place rather than
        // rebuilding it.
        var xml = OneNoteSession.Instance.GetHierarchy(
            id, OneNoteScope.HsSelf, OneNoteXmlSchema.Xs2013);

        var root = XDocument.Parse(xml).Root!;
        root.SetAttributeValue("name", newName);
        root.RemoveNodes(); // drop children; keep this node's attributes only

        OneNoteSession.Instance.UpdateHierarchy(
            root.ToString(SaveOptions.DisableFormatting), OneNoteXmlSchema.Xs2013);

        return "{\"renamed\":true}";
    }

    [McpServerTool(Name = "onenote_delete_node")]
    [Description("Deletes a section, section group, or notebook by its object ID.")]
    public static string DeleteNode(
        [Description("OneNote object ID of the node to delete.")] string id)
    {
        OneNoteSession.Instance.DeleteHierarchy(id);
        return "{\"deleted\":true}";
    }
}
