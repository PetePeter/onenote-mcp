using System.ComponentModel;
using System.Runtime.Versioning;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Read-only tools that return the content and summary metadata of a OneNote page.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class PageReadTools
{
    [McpServerTool(Name = "onenote_get_page")]
    [Description("Returns the OneNote page XML for a page ID at the requested detail level.")]
    public static string GetPage(
        [Description("OneNote object ID of the page.")] string pageId,
        [Description("Detail level: basic, selection, binaryData, fileType, or all.")] string detail) => ToolError.Guard(() =>
    {
        return OneNoteSession.Instance.GetPageContent(pageId, PageInfoMapper.MapDetail(detail), OneNoteXmlSchema.Xs2013);
    });

    [McpServerTool(Name = "onenote_get_page_info")]
    [Description("Returns summary metadata (id, title, timestamps, author, level) for a OneNote page.")]
    public static string GetPageInfo(
        [Description("OneNote object ID of the page.")] string pageId) => ToolError.Guard(() =>
    {
        var xml = OneNoteSession.Instance.GetPageContent(pageId, PageInfoMapper.MapDetail("basic"), OneNoteXmlSchema.Xs2013);
        return JsonSerializer.Serialize(PageInfoMapper.ParseMetadata(xml));
    });
}
