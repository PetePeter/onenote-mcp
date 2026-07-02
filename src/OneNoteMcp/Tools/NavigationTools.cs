using System.ComponentModel;
using System.Runtime.Versioning;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

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
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the hierarchy node (notebook/section/page).")] string hierarchyObjectId,
        [Description("Optional object ID of a content element on the page.")] string? objectId = null,
        [Description("Open in a new OneNote window.")] bool newWindow = false) =>
        ToolVersion.Guarded(version, Capability.NavigateTo, s =>
        {
            s.NavigateTo(hierarchyObjectId, objectId ?? "", newWindow);
            return "{\"navigated\":true}";
        });

    [McpServerTool(Name = "onenote_navigate_to_url")]
    [Description("Navigates the OneNote UI to a onenote: URL.")]
    public static string NavigateToUrl(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("A onenote: URL to open.")] string url,
        [Description("Open in a new OneNote window.")] bool newWindow = false) =>
        ToolVersion.Guarded(version, Capability.NavigateToUrl, s =>
        {
            s.NavigateToUrl(url, newWindow);
            return "{\"navigated\":true}";
        });

    [McpServerTool(Name = "onenote_get_hyperlink_to_object")]
    [Description("Returns a onenote: hyperlink to a hierarchy node (and optional page content object).")]
    public static string GetHyperlinkToObject(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the hierarchy node.")] string hierarchyId,
        [Description("Optional object ID of a content element on the page.")] string? pageContentObjectId = null) =>
        ToolVersion.Guarded(version, Capability.GetHyperlinkToObject, s =>
            s.GetHyperlinkToObject(hierarchyId, pageContentObjectId ?? ""));

    [McpServerTool(Name = "onenote_get_web_hyperlink_to_object")]
    [Description("Returns a web (https) hyperlink to a hierarchy node (and optional page content object).")]
    public static string GetWebHyperlinkToObject(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the hierarchy node.")] string hierarchyId,
        [Description("Optional object ID of a content element on the page.")] string? pageContentObjectId = null) =>
        ToolVersion.Guarded(version, Capability.GetWebHyperlinkToObject, s =>
            s.GetWebHyperlinkToObject(hierarchyId, pageContentObjectId ?? ""));
}
