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
/// Exports OneNote content via the Publish API: pages and sections to PDF (in
/// "single" mode the whole node is published to one PDF, in "perPage" mode each
/// page under the node is published to its own PDF in a target directory), a
/// section to a .one file, and a whole notebook to a .onepkg package.
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
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
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

        var session = ToolVersion.Route(version, Capability.Publish);

        return normalized == "single"
            ? ExportSingle(session, nodeId, outputPath)
            : ExportPerPage(session, nodeId, outputPath, interPublishDelayMs);
    }

    /// <summary>Publishes the whole node to a single PDF file.</summary>
    private static string ExportSingle(OneNoteSession session, string nodeId, string outputPath)
        => PublishNode(session, nodeId, outputPath, OneNotePublishFormat.PfPdf);

    [McpServerTool(Name = "onenote_export_one")]
    [Description("Exports a OneNote section to a .one file via the Publish API. Output is current OneNote format (2010+), not the legacy 2007 format. Returns a JSON array containing the produced file path.")]
    public static string ExportOne(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the section to export.")] string sectionId,
        [Description("Target .one file path.")] string outputPath)
    {
        var session = ToolVersion.Route(version, Capability.Publish);
        return PublishNode(session, sectionId, outputPath, OneNotePublishFormat.PfOneNote);
    }

    [McpServerTool(Name = "onenote_export_onepkg")]
    [Description("Exports a whole OneNote notebook to a .onepkg package (a Windows cabinet) via the Publish API. Output is current OneNote format (2010+), not the legacy 2007 format. Returns a JSON array containing the produced file path.")]
    public static string ExportOnepkg(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the notebook to export.")] string notebookId,
        [Description("Target .onepkg file path.")] string outputPath)
    {
        var session = ToolVersion.Route(version, Capability.Publish);
        return PublishNode(session, notebookId, outputPath, OneNotePublishFormat.PfOneNotePackage);
    }

    // Package publishes (.one/.onepkg) flush the file asynchronously — Publish can
    // return before OneNote finishes writing — so we wait for the file to appear.
    private static readonly TimeSpan PublishFileTimeout = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Publishes a single hierarchy node to one target file in the given format.
    /// Shared by the PDF, .one and .onepkg single-file exports: resolve the path,
    /// ensure the parent directory exists, clear any stale target (Publish refuses
    /// to overwrite), then publish, wait for the file to materialise, and report it.
    /// </summary>
    private static string PublishNode(OneNoteSession session, string nodeId, string outputPath, int publishFormat)
    {
        var fullPath = Path.GetFullPath(outputPath);

        var parent = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(parent))
            Directory.CreateDirectory(parent); // propagates on invalid drives

        DeleteIfExists(fullPath);
        session.Publish(nodeId, fullPath, publishFormat);
        WaitForFile(fullPath);

        return JsonSerializer.Serialize(new[] { fullPath });
    }

    /// <summary>
    /// Blocks until the published file is fully written or the timeout elapses.
    /// OneNote returns from Publish while it is still creating the file: the path
    /// first appears empty, then fills, and stays locked until the write completes.
    /// So we wait for a non-empty file whose size has stopped changing and which we
    /// can open for reading — the point at which OneNote has finished with it.
    /// </summary>
    private static void WaitForFile(string path)
    {
        var deadline = DateTime.UtcNow + PublishFileTimeout;
        long lastLength = -1;
        while (true)
        {
            if (DateTime.UtcNow >= deadline)
                throw new IOException($"OneNote did not finish producing '{path}' within {PublishFileTimeout.TotalSeconds:0}s.");

            try
            {
                var info = new FileInfo(path);
                if (info.Exists && info.Length > 0 && info.Length == lastLength)
                {
                    // Size has settled; confirm OneNote has released its write lock.
                    using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                    return;
                }
                lastLength = info.Exists ? info.Length : -1;
            }
            catch (IOException)
            {
                // Still being written/locked — keep waiting.
            }

            Thread.Sleep(150);
        }
    }

    /// <summary>Publishes each page under the node to its own PDF in a directory.</summary>
    private static string ExportPerPage(OneNoteSession session, string nodeId, string outputPath, int interPublishDelayMs)
    {
        Directory.CreateDirectory(outputPath);
        var resolvedDir = Path.GetFullPath(outputPath);

        var xml = session.GetHierarchy(nodeId, OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013);
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
            session.Publish(page.Id, fullPath, OneNotePublishFormat.PfPdf);
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
