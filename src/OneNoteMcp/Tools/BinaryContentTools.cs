using System.ComponentModel;
using System.Runtime.Versioning;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Tools that operate on individual binary/content objects within a page: fetching
/// a callback-backed binary to disk, and deleting a single content object.
/// </summary>
[McpServerToolType]
[SupportedOSPlatform("windows")]
public static class BinaryContentTools
{
    [McpServerTool(Name = "onenote_get_binary_page_content")]
    [Description("Fetches a page binary object by its callback ID, writes the decoded bytes to outputDir, and returns JSON {path, bytes}.")]
    public static string GetBinaryPageContent(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the page.")] string pageId,
        [Description("Callback ID of the binary object (from the page XML's one:CallbackID).")] string callbackId,
        [Description("Directory to write the decoded file to; created if missing.")] string outputDir) =>
        ToolVersion.Guarded(version, Capability.GetBinaryPageContent, s =>
        {
            var b64 = s.GetBinaryPageContent(pageId, callbackId);
            // Convert.FromBase64String rejects stray whitespace/newlines, so strip all.
            var stripped = string.Concat(b64.Where(c => !char.IsWhiteSpace(c)));
            var bytes = Convert.FromBase64String(stripped);

            var ext = BinaryExtractor.InferExtension(null, bytes);
            var fileName = BinaryExtractor.BuildFileName("binary", 1, ext);
            var path = BinaryExtractor.WriteBinary(outputDir, fileName, bytes);

            return JsonSerializer.Serialize(new { path, bytes = bytes.Length });
        });

    [McpServerTool(Name = "onenote_delete_page_content")]
    [Description("Deletes a single content object (by its object ID) from a page.")]
    public static string DeletePageContent(
        [Description("OneNote version token: 2007, 2010, 2013, 2016, an Office major (12/14/16), or a CLSID.")] string version,
        [Description("OneNote object ID of the page.")] string pageId,
        [Description("Object ID of the content element to delete.")] string objectId) =>
        ToolVersion.Guarded(version, Capability.DeletePageContent, s =>
        {
            s.DeletePageContent(pageId, objectId);
            return "{\"deleted\":true}";
        });
}
