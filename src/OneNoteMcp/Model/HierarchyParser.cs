using System.Xml.Linq;

namespace OneNoteMcp.Model;

/// <summary>A OneNote notebook surfaced from hierarchy XML.</summary>
public sealed record NotebookInfo(
    string Id, string Name, string Path, string LastModified, bool IsReadOnly, bool IsUnread);

/// <summary>A page matched by a search or hierarchy query, with its owning section.</summary>
public sealed record PageMatch(string Id, string Title, string SectionId);

/// <summary>
/// Pure XML parsing of OneNote hierarchy documents. Matches by element/attribute
/// LocalName so it is agnostic to the schema namespace, which varies across the
/// 2007/2010/2013 OneNote schemas. Contains no COM dependency.
/// </summary>
public static class HierarchyParser
{
    /// <summary>Extracts every notebook from a hierarchy XML document.</summary>
    public static IReadOnlyList<NotebookInfo> ParseNotebooks(string hierarchyXml)
    {
        var root = XDocument.Parse(hierarchyXml).Root;
        if (root is null) return Array.Empty<NotebookInfo>();

        return DescendantsByLocalName(root, "Notebook")
            .Select(n => new NotebookInfo(
                Id: Attr(n, "ID"),
                Name: Attr(n, "name"),
                Path: Attr(n, "path"),
                LastModified: Attr(n, "lastModifiedTime"),
                IsReadOnly: AttrBool(n, "readOnly"),
                IsUnread: AttrBool(n, "isUnread")))
            .ToList();
    }

    /// <summary>
    /// Extracts every page from a hierarchy XML document, resolving each page's
    /// section from its nearest ancestor Section element.
    /// </summary>
    public static IReadOnlyList<PageMatch> ParsePages(string hierarchyXml)
    {
        var root = XDocument.Parse(hierarchyXml).Root;
        if (root is null) return Array.Empty<PageMatch>();

        return DescendantsByLocalName(root, "Page")
            .Select(p => new PageMatch(
                Id: Attr(p, "ID"),
                Title: Attr(p, "name"),
                SectionId: NearestSectionId(p)))
            .ToList();
    }

    /// <summary>Enumerates the element and its descendants whose LocalName matches.</summary>
    private static IEnumerable<XElement> DescendantsByLocalName(XElement root, string localName) =>
        root.DescendantsAndSelf().Where(e => e.Name.LocalName == localName);

    /// <summary>Reads an attribute by local name; attributes here carry no namespace.</summary>
    private static string Attr(XElement element, string localName) =>
        element.Attribute(localName)?.Value ?? "";

    /// <summary>Parses an attribute as a boolean; absent or unrecognised values are false.</summary>
    private static bool AttrBool(XElement element, string localName)
    {
        var value = element.Attribute(localName)?.Value;
        return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "1";
    }

    /// <summary>Returns the ID of the closest ancestor Section, or "" when none.</summary>
    private static string NearestSectionId(XElement page)
    {
        var section = page.Ancestors().FirstOrDefault(a => a.Name.LocalName == "Section");
        return section is null ? "" : Attr(section, "ID");
    }
}
