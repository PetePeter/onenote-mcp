using System.ComponentModel;
using System.Runtime.Versioning;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Extracts a page's inline binary objects (images and inline files) to disk and
/// returns their absolute file paths, so clients never handle base64 payloads.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class FileExtractionTools
{
    [McpServerTool(Name = "onenote_extract_page_files")]
    [Description("Extracts a page's inline images/files to outputDir and returns JSON of the written file paths.")]
    public static string ExtractPageFiles(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the page.")] string pageId,
        [Description("Directory to write extracted files to; created if missing.")] string outputDir,
        [Description("Which binaries to extract: images, files, ink, or all.")] string filter) =>
        ToolVersion.Guarded(version, Capability.GetPageContent, s =>
        {
            var xml = s.GetPageContent(pageId, PageInfoMapper.MapDetail("all"), OneNoteXmlSchema.Xs2013);
            var pageName = PageInfoMapper.ParseMetadata(xml).Title;

            var binaries = BinaryExtractor.ExtractInlineBinaries(xml);
            var selected = ApplyFilter(binaries, filter);

            var written = new List<ExtractedFile>();
            var index = 1;
            foreach (var b in selected)
            {
                var fileName = BinaryExtractor.BuildFileName(pageName, index, b.Extension);
                // WriteBinary creates the dir and enforces the path-escape guard.
                var fullPath = BinaryExtractor.WriteBinary(outputDir, fileName, b.Bytes);
                written.Add(new ExtractedFile(
                    fullPath, b.Type, b.Width, b.Height, b.SourceElementId, b.RecognizedText));
                index++;
            }

            return JsonSerializer.Serialize(written);
        });

    /// <summary>
    /// Filters decoded binaries by the caller's selection. Empty or "all" keeps
    /// everything; "images"/"files"/"ink" keep that type. "files" excludes ink,
    /// which is its own category. Unknown filters throw.
    /// </summary>
    internal static IEnumerable<InlineBinary> ApplyFilter(
        IReadOnlyList<InlineBinary> binaries, string filter) => (filter ?? "").ToLowerInvariant() switch
    {
        "" or "all" => binaries,
        "images" => binaries.Where(b => b.Type == "image"),
        "files" => binaries.Where(b => b.Type == "file"),
        "ink" => binaries.Where(b => b.Type == "ink"),
        _ => throw new ArgumentException(
            $"Unknown filter '{filter}'. Expected one of: images, files, ink, all.", nameof(filter)),
    };
}
