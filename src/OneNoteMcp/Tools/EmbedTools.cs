using System.ComponentModel;
using System.IO;
using System.Runtime.Versioning;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Write tool that embeds a local file into a page. The server reads the file,
/// auto-routes by type (raster images inline as base64, other files attach via
/// one:InsertedFile), and appends it to the page so callers never build base64 or
/// page XML themselves. The inverse (page binary -> temp file) lives in
/// <see cref="BinaryContentTools"/> and <see cref="FileExtractionTools"/>.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class EmbedTools
{
    [McpServerTool(Name = "onenote_embed_file")]
    [Description("Embeds a local file into a page: images are inlined as base64, other files are attached. Optional x/y place the new outline; omit both to let OneNote position it. Returns JSON {kind, format, filePath}.")]
    public static string EmbedFile(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the page to embed into.")] string pageId,
        [Description("Absolute path of the local file to embed.")] string filePath,
        [Description("Optional display name for attached (non-image) files; defaults to the file name.")] string? preferredName = null,
        [Description("Optional page X coordinate for the new outline. Provide with y to place it.")] double? x = null,
        [Description("Optional page Y coordinate for the new outline. Provide with x to place it.")] double? y = null) =>
        ToolVersion.Guarded(version, Capability.UpdatePageContent, (s, major) =>
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File to embed not found: {filePath}", filePath);

            var bytes = File.ReadAllBytes(filePath);
            var embed = EmbedXmlBuilder.Build(pageId, filePath, bytes, preferredName, major, x, y);

            var schema = major == 12 ? OneNoteXmlSchema.Xs2007 : OneNoteXmlSchema.Xs2013;
            s.UpdatePageContent(embed.Xml, schema);

            return JsonSerializer.Serialize(new
            {
                kind = embed.Kind,
                format = embed.Format,
                filePath = Path.GetFullPath(filePath),
            });
        });
}
