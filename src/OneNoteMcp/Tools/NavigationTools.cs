using System.ComponentModel;
using System.Runtime.Versioning;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;

namespace OneNoteMcp.Tools;

/// <summary>
/// Tools that navigate the OneNote UI and build hyperlinks to hierarchy nodes or
/// page content objects.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class NavigationTools
{
    [McpServerTool(Name = "onenote_navigate_to")]
    [Description("Navigates the OneNote UI to a hierarchy node and optional content object.")]
    public static string NavigateTo(
        [Description("OneNote object ID of the hierarchy node (notebook/section/page).")] string hierarchyObjectId,
        [Description("Optional object ID of a content element on the page.")] string? objectId = null,
        [Description("Open in a new OneNote window.")] bool newWindow = false) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.NavigateTo(hierarchyObjectId, objectId ?? "", newWindow);
        return "{\"navigated\":true}";
    });

    [McpServerTool(Name = "onenote_navigate_to_url")]
    [Description("Navigates the OneNote UI to a onenote: URL.")]
    public static string NavigateToUrl(
        [Description("A onenote: URL to open.")] string url,
        [Description("Open in a new OneNote window.")] bool newWindow = false) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.NavigateToUrl(url, newWindow);
        return "{\"navigated\":true}";
    });

    [McpServerTool(Name = "onenote_get_hyperlink_to_object")]
    [Description("Returns a onenote: hyperlink to a hierarchy node (and optional page content object).")]
    public static string GetHyperlinkToObject(
        [Description("OneNote object ID of the hierarchy node.")] string hierarchyId,
        [Description("Optional object ID of a content element on the page.")] string? pageContentObjectId = null) => ToolError.Guard(() =>
        OneNoteSession.Instance.GetHyperlinkToObject(hierarchyId, pageContentObjectId ?? ""));

    [McpServerTool(Name = "onenote_get_web_hyperlink_to_object")]
    [Description("Returns a web (https) hyperlink to a hierarchy node (and optional page content object).")]
    public static string GetWebHyperlinkToObject(
        [Description("OneNote object ID of the hierarchy node.")] string hierarchyId,
        [Description("Optional object ID of a content element on the page.")] string? pageContentObjectId = null) => ToolError.Guard(() =>
        OneNoteSession.Instance.GetWebHyperlinkToObject(hierarchyId, pageContentObjectId ?? ""));
}
