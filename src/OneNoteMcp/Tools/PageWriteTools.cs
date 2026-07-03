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
    private const string Ns2013 =
        "http://schemas.microsoft.com/office/onenote/2013/onenote";
    private const string Ns2007 =
        "http://schemas.microsoft.com/office/onenote/2007/onenote";

    [McpServerTool(Name = "onenote_create_page")]
    [Description("Creates a new page in a section and optionally sets its title. Returns the new page's object ID.")]
    public static string CreatePage(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the section to create the page in.")] string sectionId,
        [Description("Optional page title.")] string? title = null) =>
        ToolVersion.Guarded(version, Capability.CreateNewPage, (s, major) =>
        {
            var pageId = s.CreateNewPage(sectionId);
            if (!string.IsNullOrEmpty(title))
            {
                // v12 (2007) is schema-locked to the 2007 namespace/schema; everything
                // else uses 2013. Emitting the wrong namespace makes the 2007 COM server
                // reject the title write (0x80042001), orphaning an Untitled page.
                var schema = major == 12 ? OneNoteXmlSchema.Xs2007 : OneNoteXmlSchema.Xs2013;
                try
                {
                    s.UpdatePageContent(BuildTitleXml(pageId, title, major), schema);
                }
                catch
                {
                    // Keep create+title atomic: a failed title write must not leave an
                    // orphan page behind. Best-effort delete, then surface the original error.
                    try { s.DeleteHierarchy(pageId); } catch { /* best-effort cleanup */ }
                    throw;
                }
            }
            return pageId;
        });

    [McpServerTool(Name = "onenote_update_page")]
    [Description("Writes full OneNote page XML back to a page, applying the changes it carries.")]
    public static string UpdatePage(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("Full OneNote page XML (root <one:Page> carrying the page ID) to write.")] string pageXml) =>
        ToolVersion.Guarded(version, Capability.UpdatePageContent, s =>
        {
            PageXmlValidator.Validate(pageXml);
            s.UpdatePageContent(pageXml, OneNoteXmlSchema.Xs2013);
            return "{\"updated\":true}";
        });

    [McpServerTool(Name = "onenote_delete_page")]
    [Description("Deletes a OneNote page by its object ID.")]
    public static string DeletePage(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the page to delete.")] string pageId) =>
        ToolVersion.Guarded(version, Capability.DeleteHierarchy, s =>
        {
            s.DeleteHierarchy(pageId);
            return "{\"deleted\":true}";
        });

    /// <summary>
    /// Builds minimal page XML that sets just the title for a page ID. Built with
    /// System.Xml.Linq so the title text is auto-escaped against XML injection.
    /// </summary>
    internal static string BuildTitleXml(string pageId, string title, int major)
    {
        XNamespace one = major == 12 ? Ns2007 : Ns2013;
        var page = new XElement(one + "Page",
            new XAttribute(XNamespace.Xmlns + "one", one.NamespaceName),
            new XAttribute("ID", pageId),
            new XElement(one + "Title",
                new XElement(one + "OE",
                    new XElement(one + "T", title))));

        return page.ToString(SaveOptions.DisableFormatting);
    }
}
