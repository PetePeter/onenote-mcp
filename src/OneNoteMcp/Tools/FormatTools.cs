using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Xml.Linq;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Format-awareness tools: report whether a .one/.onetoc2 file is the current 2010+
/// or legacy 2007 format (a pure header read), and best-effort upgrade a section to
/// the current format by republishing it. Both tools return JSON and, deliberately,
/// never throw — a bad input yields an honest failure report rather than crashing the
/// server, because callers batch these across many sections OneNote may not open.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class FormatTools
{
    [McpServerTool(Name = "onenote_detect_format")]
    [Description("Detects whether a OneNote .one/.onetoc2 file is the current 2010+ format or the legacy 2007 format by reading its file header. Accepts a file path or a OneNote object ID (resolved to its backing file path). Read-only: never modifies the file. Returns a JSON report.")]
    public static string DetectFormat(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("A file path to a .one/.onetoc2 file, or a OneNote object ID resolved to its backing file.")] string pathOrNodeId)
    {
        var path = ResolveToFilePath(pathOrNodeId, TryGetSession(version));
        if (path is null)
            return Report(pathOrNodeId, new OneFileFormatResult(
                OneNoteFileFormat.NotOneNoteFile,
                "unknown",
                "Input was neither an existing file nor a OneNote node whose backing file exists on disk."));

        return Report(pathOrNodeId, OneFileFormatSniffer.SniffFile(path));
    }

    [McpServerTool(Name = "onenote_convert_section")]
    [Description("Best-effort upgrade of a OneNote section to the current 2010+ .one format via the Publish API. Opens the section if OneNote accepts it and republishes to outputPath. Reports failure per-section (never crashes) for sections OneNote cannot open. Returns a JSON report.")]
    public static string ConvertSection(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the section to upgrade.")] string sectionId,
        [Description("Target .one file path for the republished current-format section.")] string outputPath)
    {
        try
        {
            // Reuse ExportOne so the publish+wait-for-file logic lives in one place;
            // it publishes PfOneNote (current format) and returns the resolved path.
            var exported = ExportTools.ExportOne(version, sectionId, outputPath);
            var producedPath = FirstPath(exported) ?? Path.GetFullPath(outputPath);

            // Sniff the produced file to confirm and report the resulting format.
            var sniff = OneFileFormatSniffer.SniffFile(producedPath);
            return JsonSerializer.Serialize(new
            {
                success = true,
                sectionId,
                outputPath = producedPath,
                format = FormatString(sniff.Format),
                detail = sniff.Detail,
            });
        }
        catch (Exception ex)
        {
            // Per-section failure must never crash a batch conversion.
            return JsonSerializer.Serialize(new
            {
                success = false,
                sectionId,
                outputPath,
                reason = ComErrorMapper.Describe(ex),
            });
        }
    }

    /// <summary>
    /// Tries to create a session for the given version, returning null when the
    /// version is invalid or resolve fails. Used by DetectFormat which must never
    /// throw.
    /// </summary>
    private static OneNoteSession? TryGetSession(string version)
    {
        try
        {
            var (clsid, _) = ToolVersion.Resolve(version);
            return OneNoteSession.For(clsid);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Resolves the input to a readable file path: a direct file wins; otherwise we
    /// try to treat it as a OneNote object ID and read its backing "path" attribute
    /// via the session. Returns null when nothing resolves to an existing file.
    /// </summary>
    private static string? ResolveToFilePath(string pathOrNodeId, OneNoteSession? session)
    {
        if (File.Exists(pathOrNodeId))
            return pathOrNodeId;

        if (session is null)
            return null;

        try
        {
            var xml = session.GetHierarchy(pathOrNodeId, OneNoteScope.HsSelf, OneNoteXmlSchema.Xs2013);
            var root = XDocument.Parse(xml).Root;
            var backing = root?.Attributes()
                .FirstOrDefault(a => string.Equals(a.Name.LocalName, "path", StringComparison.Ordinal))?.Value;

            if (!string.IsNullOrEmpty(backing) && File.Exists(backing))
                return backing;
        }
        catch
        {
            // Not a resolvable node (bad id, COM unavailable, malformed XML): fall
            // through to the not-found report rather than surfacing an exception.
        }

        return null;
    }

    /// <summary>Serialises a sniff result into the tool's camelCase JSON report shape.</summary>
    private static string Report(string input, OneFileFormatResult result)
        => JsonSerializer.Serialize(new
        {
            input,
            format = FormatString(result.Format),
            fileType = result.FileType,
            is2007 = result.Is2007,
            isCurrent = result.IsCurrent,
            detail = result.Detail,
        });

    /// <summary>Maps the format enum to the tool's stable camelCase wire strings.</summary>
    private static string FormatString(OneNoteFileFormat format) => format switch
    {
        OneNoteFileFormat.Current2010Plus => "current2010Plus",
        OneNoteFileFormat.Legacy2007 => "legacy2007",
        _ => "notOneNoteFile",
    };

    /// <summary>Extracts the first path from ExportOne's JSON array result, if any.</summary>
    private static string? FirstPath(string exportedJson)
    {
        using var doc = JsonDocument.Parse(exportedJson);
        if (doc.RootElement.ValueKind == JsonValueKind.Array && doc.RootElement.GetArrayLength() > 0)
            return doc.RootElement[0].GetString();
        return null;
    }
}
