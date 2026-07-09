using System.IO;
using System.Xml.Linq;

namespace OneNoteMcp.Model;

/// <summary>
/// Pure, COM-free builder for the page-update XML that embeds a local file into a
/// page. Auto-routes by type: raster images (png/jpg/gif/bmp) are inlined as base64
/// inside <c>one:Image</c>; every other file is attached natively via
/// <c>one:InsertedFile</c> (which references the file path on disk — OneNote copies
/// it in, so no base64 is used). The embedded object is wrapped in a fresh
/// <c>one:Outline</c> carrying the page ID so <c>UpdatePageContent</c> appends it.
/// Built with <see cref="XElement"/> so all attribute/text values are auto-escaped.
/// </summary>
public static class EmbedXmlBuilder
{
    private const string Ns2013 =
        "http://schemas.microsoft.com/office/onenote/2013/onenote";
    private const string Ns2007 =
        "http://schemas.microsoft.com/office/onenote/2007/onenote";

    private static readonly HashSet<string> ImageExtensions =
        new(StringComparer.OrdinalIgnoreCase) { "png", "jpg", "jpeg", "gif", "bmp" };

    /// <summary>The kind of embed produced, surfaced back to the caller as JSON.</summary>
    public sealed record EmbedResult(string Xml, string Kind, string Format);

    /// <summary>
    /// Builds the append XML for <paramref name="filePath"/> (already read into
    /// <paramref name="bytes"/>). <paramref name="major"/> selects the 2007 vs 2013
    /// namespace. When both <paramref name="positionX"/> and <paramref name="positionY"/>
    /// are supplied, the new outline is placed at that page coordinate; otherwise
    /// OneNote chooses the position.
    /// </summary>
    public static EmbedResult Build(
        string pageId, string filePath, byte[] bytes, string? preferredName,
        int major, double? positionX, double? positionY)
    {
        XNamespace one = major == 12 ? Ns2007 : Ns2013;

        var fileExt = Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();
        var sniffExt = BinaryExtractor.InferExtension(null, bytes);
        var format = string.IsNullOrEmpty(fileExt) ? sniffExt : fileExt;
        var isImage = ImageExtensions.Contains(format) || ImageExtensions.Contains(sniffExt);

        XElement obj;
        string kind;
        if (isImage)
        {
            obj = new XElement(one + "Image",
                new XAttribute("format", format),
                new XElement(one + "Data", Convert.ToBase64String(bytes)));
            kind = "image";
        }
        else
        {
            var name = string.IsNullOrWhiteSpace(preferredName)
                ? Path.GetFileName(filePath)
                : preferredName;
            obj = new XElement(one + "InsertedFile",
                new XAttribute("pathCache", Path.GetFullPath(filePath)),
                new XAttribute("preferredName", name));
            kind = "file";
        }

        var outline = new XElement(one + "Outline");
        if (positionX is { } x && positionY is { } y)
            outline.Add(new XElement(one + "Position",
                new XAttribute("x", x.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new XAttribute("y", y.ToString(System.Globalization.CultureInfo.InvariantCulture))));
        outline.Add(new XElement(one + "OEChildren",
            new XElement(one + "OE", obj)));

        var page = new XElement(one + "Page",
            new XAttribute(XNamespace.Xmlns + "one", one.NamespaceName),
            new XAttribute("ID", pageId),
            outline);

        return new EmbedResult(page.ToString(SaveOptions.DisableFormatting), kind, format);
    }
}
