using System.ComponentModel;
using System.Runtime.Versioning;
using System.Xml.Linq;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Write tools that create, update, and delete OneNote pages. Per the project's
/// "no guardrails on destructive ops" decision these apply changes without a
/// conflict check; update payloads are validated for shape before the COM call.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class PageWriteTools
{
    private static readonly XNamespace One =
        "http://schemas.microsoft.com/office/onenote/2013/onenote";

    [McpServerTool(Name = "onenote_create_page")]
    [Description("Creates a new page in a section and optionally sets its title. Returns the new page's object ID.")]
    public static string CreatePage(
        [Description("OneNote object ID of the section to create the page in.")] string sectionId,
        [Description("Optional page title.")] string? title = null) => ToolError.Guard(() =>
    {
        var pageId = OneNoteSession.Instance.CreateNewPage(sectionId);

        if (!string.IsNullOrEmpty(title))
            OneNoteSession.Instance.UpdatePageContent(BuildTitleXml(pageId, title), OneNoteXmlSchema.Xs2013);

        return pageId;
    });

    [McpServerTool(Name = "onenote_update_page")]
    [Description("Writes full OneNote page XML back to a page, applying the changes it carries.")]
    public static string UpdatePage(
        [Description("Full OneNote page XML (root <one:Page> carrying the page ID) to write.")] string pageXml) => ToolError.Guard(() =>
    {
        PageXmlValidator.Validate(pageXml);
        OneNoteSession.Instance.UpdatePageContent(pageXml, OneNoteXmlSchema.Xs2013);
        return "{\"updated\":true}";
    });

    [McpServerTool(Name = "onenote_delete_page")]
    [Description("Deletes a OneNote page by its object ID.")]
    public static string DeletePage(
        [Description("OneNote object ID of the page to delete.")] string pageId) => ToolError.Guard(() =>
    {
        OneNoteSession.Instance.DeleteHierarchy(pageId);
        return "{\"deleted\":true}";
    });

    /// <summary>
    /// Builds minimal page XML that sets just the title for a page ID. Built with
    /// System.Xml.Linq so the title text is auto-escaped against XML injection.
    /// </summary>
    private static string BuildTitleXml(string pageId, string title)
    {
        var page = new XElement(One + "Page",
            new XAttribute(XNamespace.Xmlns + "one", One.NamespaceName),
            new XAttribute("ID", pageId),
            new XElement(One + "Title",
                new XElement(One + "OE",
                    new XElement(One + "T", title))));

        return page.ToString(SaveOptions.DisableFormatting);
    }
}
