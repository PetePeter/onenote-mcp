using System.Xml.Linq;

namespace OneNoteMcp.Model;

/// <summary>Summary metadata extracted from a OneNote page's XML.</summary>
public sealed record PageMetadata(
    string Id, string Title, string DateTime, string LastModified, string Author, string Level);

/// <summary>
/// Pure, COM-free mapping and parsing for the page read tools: translating a
/// caller detail level to OneNote PageInfo flags, and pulling summary metadata
/// out of page XML. Matches by element/attribute LocalName so it is agnostic to
/// the schema namespace, which varies across the OneNote 2007/2010/2013 schemas.
/// </summary>
public static class PageInfoMapper
{
    /// <summary>
    /// Maps a caller-supplied detail level to its OneNote PageInfo flag value.
    /// Null or empty defaults to "basic" (0).
    /// </summary>
    public static int MapDetail(string detail) => (detail ?? "").ToLowerInvariant() switch
    {
        "" or "basic" => 0,
        "selection"   => 1,
        "binarydata"  => 2,
        "filetype"    => 4,
        "all"         => 7,
        _ => throw new ArgumentException(
            $"Unknown detail '{detail}'. Expected one of: basic, selection, binaryData, fileType, all.",
            nameof(detail)),
    };

    /// <summary>
    /// Extracts summary metadata from a OneNote page XML document. The title
    /// prefers the first Title/T text run (CDATA unwrapped) and falls back to the
    /// Page element's name attribute. Missing attributes default to "".
    /// </summary>
    public static PageMetadata ParseMetadata(string pageXml)
    {
        var root = XDocument.Parse(pageXml).Root
            ?? throw new ArgumentException("Page XML has no root element.", nameof(pageXml));

        return new PageMetadata(
            Id: Attr(root, "ID"),
            Title: Title(root),
            DateTime: Attr(root, "dateTime"),
            LastModified: Attr(root, "lastModifiedTime"),
            Author: Attr(root, "author"),
            Level: Attr(root, "pageLevel"));
    }

    /// <summary>Reads an attribute by local name; attributes here carry no namespace.</summary>
    private static string Attr(XElement element, string localName) =>
        element.Attribute(localName)?.Value ?? "";

    /// <summary>
    /// Returns the trimmed text of the first Title/T run, or the Page name
    /// attribute when no title text is present.
    /// </summary>
    private static string Title(XElement page)
    {
        var title = page.Descendants().FirstOrDefault(e => e.Name.LocalName == "Title");
        var text = title?.Descendants().FirstOrDefault(e => e.Name.LocalName == "T")?.Value;
        return (string.IsNullOrEmpty(text) ? Attr(page, "name") : text).Trim();
    }
}
