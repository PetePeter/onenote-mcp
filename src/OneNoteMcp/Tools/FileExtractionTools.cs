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
        [Description("OneNote object ID of the page.")] string pageId,
        [Description("Directory to write extracted files to; created if missing.")] string outputDir,
        [Description("Which binaries to extract: images, files, or all.")] string filter) => ToolError.Guard(() =>
    {
        var xml = OneNoteSession.Instance.GetPageContent(
            pageId, PageInfoMapper.MapDetail("all"), OneNoteXmlSchema.Xs2013);
        var pageName = PageInfoMapper.ParseMetadata(xml).Title;

        var binaries = BinaryExtractor.ExtractInlineBinaries(xml);
        var selected = ApplyFilter(binaries, filter);

        Directory.CreateDirectory(outputDir);
        var resolvedDir = Path.GetFullPath(outputDir);

        var written = new List<ExtractedFile>();
        var index = 1;
        foreach (var b in selected)
        {
            var fileName = BinaryExtractor.BuildFileName(pageName, index, b.Extension);
            var fullPath = Path.GetFullPath(Path.Combine(resolvedDir, fileName));
            // Defence in depth: never write outside the caller's chosen directory.
            if (!fullPath.StartsWith(resolvedDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Derived path '{fullPath}' escapes outputDir.");
            File.WriteAllBytes(fullPath, b.Bytes);
            written.Add(new ExtractedFile(fullPath, b.Type, b.Width, b.Height, b.SourceElementId));
            index++;
        }

        return JsonSerializer.Serialize(written);
    });

    /// <summary>
    /// Filters decoded binaries by the caller's selection. Empty or "all" keeps
    /// everything; "images"/"files" keep that type. Unknown filters throw.
    /// </summary>
    private static IEnumerable<InlineBinary> ApplyFilter(
        IReadOnlyList<InlineBinary> binaries, string filter) => (filter ?? "").ToLowerInvariant() switch
    {
        "" or "all" => binaries,
        "images" => binaries.Where(b => b.Type == "image"),
        "files" => binaries.Where(b => b.Type == "file"),
        _ => throw new ArgumentException(
            $"Unknown filter '{filter}'. Expected one of: images, files, all.", nameof(filter)),
    };
}
