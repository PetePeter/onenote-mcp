using System.ComponentModel;
using System.IO;
using System.Runtime.Versioning;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Write tools that open, create, and close OneNote notebooks. A notebook is a
/// folder on disk; OneNote loads it by folder path via OpenHierarchy.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class NotebookTools
{
    [McpServerTool(Name = "onenote_open_notebook")]
    [Description("Opens an existing notebook folder in OneNote. Returns the notebook's object ID.")]
    public static string OpenNotebook(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Filesystem path to the notebook folder.")] string path) =>
        ToolVersion.Guarded(version, Capability.OpenHierarchy, s =>
            s.OpenHierarchy(NormalizeNotebookPath(path), "", OneNoteCreateFileType.CftNone));

    [McpServerTool(Name = "onenote_create_notebook")]
    [Description("Creates a new notebook at the given folder path. Returns the notebook's object ID.")]
    public static string CreateNotebook(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Filesystem path for the new notebook folder.")] string path) =>
        ToolVersion.Guarded(version, Capability.OpenHierarchy, s =>
            s.OpenHierarchy(NormalizeNotebookPath(path), "", OneNoteCreateFileType.CftNotebook));

    [McpServerTool(Name = "onenote_close_notebook")]
    [Description("Closes an open notebook. Does not delete files on disk.")]
    public static string CloseNotebook(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the notebook to close.")] string notebookId) =>
        ToolVersion.Guarded(version, Capability.CloseNotebook, s =>
        {
            s.CloseNotebook(notebookId);
            return "{\"closed\":true}";
        });

    /// <summary>
    /// Ensures the notebook folder path ends with a directory separator. OneNote
    /// silently fails to load a notebook whose folder name contains a dot unless
    /// the path is terminated with a separator, so we always append one.
    /// </summary>
    private static string NormalizeNotebookPath(string path)
    {
        if (path.EndsWith(Path.DirectorySeparatorChar) ||
            path.EndsWith(Path.AltDirectorySeparatorChar))
            return path;

        return path + Path.DirectorySeparatorChar;
    }
}
