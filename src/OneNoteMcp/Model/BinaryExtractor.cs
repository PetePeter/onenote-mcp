using System.IO;
using System.Xml.Linq;

namespace OneNoteMcp.Model;

/// <summary>A decoded inline binary from a page, ready to be written to disk.</summary>
public sealed record InlineBinary(
    byte[] Bytes, string Type, string Extension, int? Width, int? Height, string SourceElementId,
    string? RecognizedText = null);

/// <summary>A record describing a binary that was extracted to a file on disk.</summary>
public sealed record ExtractedFile(
    string Path, string Type, int? Width, int? Height, string SourceElementId,
    string? RecognizedText = null);

/// <summary>
/// Pure, COM-free extraction of inline binary objects from OneNote page XML.
/// Handles <c>one:Image</c> elements (each wrapping a base64 <c>one:Data</c> run)
/// and any other element carrying an inline <c>one:Data</c> child. Namespace-agnostic
/// (matches by LocalName) so it works across the 2007/2010/2013 schemas.
/// Attachment extraction that requires GetBinaryPageContent (callback-backed
/// <c>one:InsertedFile</c>) is out of scope here — only inline data is decoded.
/// </summary>
public static class BinaryExtractor
{
    /// <summary>
    /// Parses page XML and returns every inline binary object, decoded to bytes.
    /// Elements with a <c>format</c>/local name of "Image" are typed "image";
    /// all other inline-data carriers are typed "file".
    /// </summary>
    public static IReadOnlyList<InlineBinary> ExtractInlineBinaries(string pageXml)
    {
        var root = XDocument.Parse(pageXml).Root
            ?? throw new ArgumentException("Page XML has no root element.", nameof(pageXml));

        var result = new List<InlineBinary>();

        // A carrier is any element with a non-empty Data child; images wrap their
        // pixels in one:Data, other inline-file objects do the same.
        foreach (var carrier in root.DescendantsAndSelf())
        {
            var data = carrier.Elements().FirstOrDefault(e => e.Name.LocalName == "Data");
            if (data is null || string.IsNullOrWhiteSpace(data.Value))
                continue;

            // Convert.FromBase64String rejects stray whitespace/newlines, so strip all.
            var stripped = string.Concat(data.Value.Where(c => !char.IsWhiteSpace(c)));
            var bytes = Convert.FromBase64String(stripped);

            // Ink carriers inline their ISF (Ink Serialized Format) as base64 Data
            // at detail=all. They must be typed "ink" with an explicit "isf" extension
            // rather than sniffed like images/generic files.
            var local = carrier.Name.LocalName;
            var isInk = local is "InkDrawing" or "InkWord" or "InkParagraph";

            var type = isInk ? "ink" : local == "Image" ? "image" : "file";
            var format = carrier.Attribute("format")?.Value;
            var extension = isInk ? "isf" : InferExtension(format, bytes);

            // Geometry lives in a Size child (InkDrawing/Image); InkWord instead carries
            // width/height as its own attributes, so fall back to those when no child.
            var size = carrier.Elements().FirstOrDefault(e => e.Name.LocalName == "Size");
            var width = ParseDimension(size?.Attribute("width")?.Value)
                ?? ParseDimension(carrier.Attribute("width")?.Value);
            var height = ParseDimension(size?.Attribute("height")?.Value)
                ?? ParseDimension(carrier.Attribute("height")?.Value);

            var sourceId = carrier.Attribute("objectID")?.Value ?? "";
            var recognizedText = carrier.Attribute("recognizedText")?.Value;

            result.Add(new InlineBinary(
                bytes, type, extension, width, height, sourceId, recognizedText));
        }

        return result;
    }

    /// <summary>
    /// Parses a OneNote size attribute (e.g. "48" or "48.0") to an int, returning
    /// null when absent or unparseable.
    /// </summary>
    private static int? ParseDimension(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (double.TryParse(value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
            return (int)d;
        return null;
    }

    /// <summary>
    /// Infers a file extension (no leading dot) from a OneNote format attribute
    /// when present, else by sniffing the byte magic header, else "bin".
    /// </summary>
    public static string InferExtension(string? format, byte[] bytes)
    {
        if (!string.IsNullOrWhiteSpace(format))
            return format.Trim().TrimStart('.').ToLowerInvariant();

        if (StartsWith(bytes, 0x89, 0x50, 0x4E, 0x47)) return "png";
        if (StartsWith(bytes, 0xFF, 0xD8, 0xFF)) return "jpg";
        if (StartsWith(bytes, 0x47, 0x49, 0x46, 0x38)) return "gif";
        if (StartsWith(bytes, 0x42, 0x4D)) return "bmp";
        if (StartsWith(bytes, 0x25, 0x50, 0x44, 0x46)) return "pdf";

        return "bin";
    }

    /// <summary>Returns true when <paramref name="bytes"/> begins with the given magic bytes.</summary>
    private static bool StartsWith(byte[] bytes, params byte[] magic)
    {
        if (bytes.Length < magic.Length)
            return false;
        for (var i = 0; i < magic.Length; i++)
            if (bytes[i] != magic[i])
                return false;
        return true;
    }

    /// <summary>
    /// Builds a filesystem-safe file name "{pageName}_{index}.{ext}", replacing
    /// characters invalid on the platform with underscores. Index is 1-based.
    /// </summary>
    public static string BuildFileName(string pageName, int index, string extension)
    {
        var name = string.IsNullOrEmpty(pageName) ? "page" : pageName;
        var safe = Sanitize(name);
        // Extension comes from an untrusted page attribute; sanitize it too so a
        // crafted format like "../../x" cannot inject path separators into the name.
        var ext = Sanitize((extension ?? "").TrimStart('.'));
        return $"{safe}_{index}.{ext}";
    }

    /// <summary>
    /// Resolves the caller's output directory: returns it when supplied, else a
    /// fresh unique directory under the system temp path (created here) so callers
    /// that just want "grab it to a file" need not choose a location.
    /// </summary>
    public static string ResolveOutputDir(string? outputDir)
    {
        if (!string.IsNullOrWhiteSpace(outputDir))
            return outputDir;
        var temp = Path.Combine(Path.GetTempPath(), "onenote-mcp", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        return temp;
    }

    /// <summary>Replaces every character invalid in a file name with an underscore.</summary>
    private static string Sanitize(string value) => new(value
        .Select(c => Array.IndexOf(Path.GetInvalidFileNameChars(), c) >= 0 ? '_' : c)
        .ToArray());

    /// <summary>
    /// Writes <paramref name="bytes"/> to <paramref name="fileName"/> inside
    /// <paramref name="outputDir"/> (created if missing) and returns the absolute
    /// path. Defence in depth: throws if the resolved path escapes outputDir, so a
    /// crafted file name can never write outside the caller's chosen directory.
    /// </summary>
    public static string WriteBinary(string outputDir, string fileName, byte[] bytes)
    {
        Directory.CreateDirectory(outputDir);
        var resolvedDir = Path.GetFullPath(outputDir);
        var fullPath = Path.GetFullPath(Path.Combine(resolvedDir, fileName));
        if (!fullPath.StartsWith(resolvedDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Derived path '{fullPath}' escapes outputDir.");
        File.WriteAllBytes(fullPath, bytes);
        return fullPath;
    }
}
