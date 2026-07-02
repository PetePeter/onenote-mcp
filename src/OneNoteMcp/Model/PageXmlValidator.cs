using System.Xml;
using System.Xml.Linq;

namespace OneNoteMcp.Model;

/// <summary>
/// Pure, COM-free guard that runs before any UpdatePageContent COM call. It rejects
/// page XML that OneNote would refuse (or silently mangle) and turns each failure
/// into an actionable ArgumentException so callers can fix their payload. Matches by
/// element LocalName like PageInfoMapper.cs, but pins the OneNote schema namespace so
/// foreign or missing namespaces are caught early.
/// </summary>
public static class PageXmlValidator
{
    // Any OneNote schema namespace (2007/2010/2013) contains this stable fragment.
    private const string OneNoteNsFragment = "schemas.microsoft.com/office/onenote";

    /// <summary>
    /// Throws <see cref="ArgumentException"/> with an actionable message when the XML
    /// is not a well-formed OneNote page rooted in a &lt;Page&gt; element in a onenote
    /// schema namespace. Returns silently when valid.
    /// </summary>
    public static void Validate(string pageXml)
    {
        if (string.IsNullOrWhiteSpace(pageXml))
            throw new ArgumentException("Page XML is empty. Provide the full <one:Page> XML to write.", nameof(pageXml));

        XDocument doc;
        try
        {
            doc = XDocument.Parse(pageXml);
        }
        catch (XmlException ex)
        {
            throw new ArgumentException(
                $"Page XML is not well-formed: {ex.Message}", nameof(pageXml), ex);
        }

        var root = doc.Root!;

        if (!string.Equals(root.Name.LocalName, "Page", StringComparison.Ordinal))
            throw new ArgumentException(
                $"Root element must be <Page>, but was <{root.Name.LocalName}>.", nameof(pageXml));

        var ns = root.Name.NamespaceName;
        if (ns.IndexOf(OneNoteNsFragment, StringComparison.OrdinalIgnoreCase) < 0)
            throw new ArgumentException(
                $"Root <Page> must be in a OneNote schema namespace (containing '{OneNoteNsFragment}'), " +
                $"but was '{ns}'.", nameof(pageXml));

        // Without an ID, UpdatePageContent silently CREATES a new page instead of
        // updating the intended one — a hard-to-spot data divergence. Require it.
        if (string.IsNullOrEmpty(root.Attribute("ID")?.Value))
            throw new ArgumentException(
                "Root <Page> must carry an 'ID' attribute so the update targets the correct page.",
                nameof(pageXml));
    }
}
