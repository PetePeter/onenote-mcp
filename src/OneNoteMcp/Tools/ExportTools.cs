using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Threading;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Exports OneNote pages and sections to PDF via the Publish API. In "single"
/// mode the whole node is published to one PDF; in "perPage" mode each page under
/// the node is published to its own PDF in a target directory.
/// Two OneNote gotchas shape this code: Publish fails when the target file already
/// exists (so we delete first), and OneNote can wedge when publishes are hammered
/// (so perPage mode sleeps between successive publishes).
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class ExportTools
{
    // Keep derived filenames well under path limits while preserving readability.
    private const int MaxTitleLength = 80;

    [McpServerTool(Name = "onenote_export_pdf")]
    [Description("Exports a OneNote page or section to PDF via the Publish API. mode 'single' writes one PDF to outputPath (a file path); mode 'perPage' writes one PDF per page into outputPath (a directory). Returns a JSON array of the produced file paths.")]
    public static string ExportPdf(
        [Description("OneNote object ID of the page or section to export.")] string nodeId,
        [Description("For mode 'single': the target .pdf file path. For mode 'perPage': the target directory (created if missing).")] string outputPath,
        [Description("'single' (one PDF for the whole node) or 'perPage' (one PDF per page under the node).")] string mode = "single",
        [Description("Milliseconds to wait between page publishes in perPage mode; OneNote can wedge when publishes are hammered.")] int interPublishDelayMs = 2000)
    {
        // Validate the mode before any filesystem or COM side effect so a bad
        // request never leaves scratch files behind.
        var normalized = (mode ?? "").ToLowerInvariant();
        if (normalized != "single" && normalized != "perpage")
            throw new ArgumentException($"Unknown mode '{mode}'. Expected 'single' or 'perPage'.", nameof(mode));

        return normalized == "single"
            ? ExportSingle(nodeId, outputPath)
            : ExportPerPage(nodeId, outputPath, interPublishDelayMs);
    }

    /// <summary>Publishes the whole node to a single PDF file.</summary>
    private static string ExportSingle(string nodeId, string outputPath)
    {
        var fullPath = Path.GetFullPath(outputPath);

        var parent = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(parent))
            Directory.CreateDirectory(parent); // propagates on invalid drives

        DeleteIfExists(fullPath);
        OneNoteSession.Instance.Publish(nodeId, fullPath, OneNotePublishFormat.PfPdf);

        return JsonSerializer.Serialize(new[] { fullPath });
    }

    /// <summary>Publishes each page under the node to its own PDF in a directory.</summary>
    private static string ExportPerPage(string nodeId, string outputPath, int interPublishDelayMs)
    {
        Directory.CreateDirectory(outputPath);
        var resolvedDir = Path.GetFullPath(outputPath);

        var xml = OneNoteSession.Instance.GetHierarchy(nodeId, OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013);
        var pages = HierarchyParser.ParsePages(xml);
        if (pages.Count == 0)
            return JsonSerializer.Serialize(Array.Empty<string>());

        var written = new List<string>();
        var index = 1;
        foreach (var page in pages)
        {
            // Space out publishes (skipping the first) so OneNote doesn't wedge.
            if (index > 1 && interPublishDelayMs > 0)
                Thread.Sleep(interPublishDelayMs);

            var fileName = $"{index:D3} {SanitizeTitle(page.Title)}.pdf";
            var fullPath = Path.GetFullPath(Path.Combine(resolvedDir, fileName));
            // Defence in depth: never write outside the caller's chosen directory.
            if (!fullPath.StartsWith(resolvedDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Derived path '{fullPath}' escapes outputPath.");

            DeleteIfExists(fullPath);
            OneNoteSession.Instance.Publish(page.Id, fullPath, OneNotePublishFormat.PfPdf);
            written.Add(fullPath);
            index++;
        }

        return JsonSerializer.Serialize(written);
    }

    /// <summary>
    /// Publish refuses to overwrite, so remove any stale target beforehand.
    /// </summary>
    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    /// <summary>
    /// Turns a page title into a safe, bounded filename component, replacing
    /// invalid characters and falling back to "Untitled" for empty titles.
    /// </summary>
    private static string SanitizeTitle(string title)
    {
        var cleaned = string.IsNullOrWhiteSpace(title) ? "Untitled" : title.Trim();
        foreach (var c in Path.GetInvalidFileNameChars())
            cleaned = cleaned.Replace(c, '_');

        cleaned = cleaned.Trim();
        if (cleaned.Length == 0)
            cleaned = "Untitled";
        if (cleaned.Length > MaxTitleLength)
            cleaned = cleaned.Substring(0, MaxTitleLength).Trim();

        return cleaned;
    }
}
