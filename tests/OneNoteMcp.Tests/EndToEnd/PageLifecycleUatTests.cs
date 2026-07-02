using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests.EndToEnd;

/// <summary>
/// User-acceptance-level end-to-end test. Where every other integration test
/// exercises a single tool in isolation, this one drives the ACTUAL shipped
/// [McpServerTool] facade methods (Tools/*, never OneNoteSession directly) as a
/// real MCP client would: one continuous section + page lifecycle where each step
/// reads back the previous step's write.
///
/// Lifecycle: CreateSection -> CreatePage -> GetPageInfo (title) -> UpdatePage
/// (mutate a text run) -> GetPage all (mutation present + title preserved) ->
/// ExtractPageFiles on a page seeded with a known PNG (magic bytes) -> RenameNode
/// (hierarchy reflects) -> DeletePage -> DeleteNode.
///
/// It creates its OWN scratch section + pages under the shared fixture notebook
/// and tears them down in a finally, so it never mutates content other tests
/// assert against and leaves %TEMP% clean. It early-returns (skips) when the
/// fixture could not be built over COM.
/// </summary>
[Collection("OneNote COM")]
public sealed class PageLifecycleUatTests
{
    private const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    // Clean-room 1x1 transparent PNG (generated, not copied from any source notebook).
    private const string OnePixelPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    private static readonly byte[] PngMagic = { 0x89, 0x50, 0x4E, 0x47 };

    private static readonly JsonSerializerOptions JsonOpts =
        new() { PropertyNameCaseInsensitive = true };

    private readonly FixtureNotebook _fx;

    public PageLifecycleUatTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void FullLifecycle_ThroughMcpToolFacades_RoundTripsEachStep()
    {
        if (!_fx.Available) return;

        var suffix = Guid.NewGuid().ToString("N");
        var sectionName = "UAT Section " + suffix;
        var pageTitle = "UAT Page " + suffix;

        string? sectionId = null;
        string? textPageId = null;
        string? imagePageId = null;
        var outDir = Path.Combine(
            Path.GetTempPath(), "onenote-mcp-tests", "uat-" + suffix);

        try
        {
            // 1. CreateSection -> the new section appears in the notebook hierarchy.
            sectionId = SectionTools.CreateSection(FixtureNotebook.TestVersion, _fx.NotebookId, sectionName);
            Assert.False(string.IsNullOrEmpty(sectionId), "CreateSection must return an object ID");
            Assert.Contains(NotebookSections(), s => s.Id == sectionId && s.Name == sectionName);

            // 2. CreatePage with a title -> the new page appears under the section.
            textPageId = PageWriteTools.CreatePage(FixtureNotebook.TestVersion, sectionId, pageTitle);
            Assert.False(string.IsNullOrEmpty(textPageId), "CreatePage must return an object ID");
            Assert.Contains(SectionPages(sectionId), p => p.Id == textPageId);

            // 3. GetPageInfo reads back step 2's write: the title matches.
            var infoJson = PageReadTools.GetPageInfo(FixtureNotebook.TestVersion, textPageId);
            var info = JsonSerializer.Deserialize<PageMetadata>(infoJson, JsonOpts)!;
            Assert.Equal(pageTitle, info.Title);

            // 4. UpdatePage mutates one body text run, keeping the title. Seed an
            //    "original" run first, read it back, then mutate it to a new value.
            PageWriteTools.UpdatePage(FixtureNotebook.TestVersion, BuildPageXml(textPageId, pageTitle, "uat-run-original"));

            var seeded = PageReadTools.GetPage(FixtureNotebook.TestVersion, textPageId, "all");
            Assert.Contains("uat-run-original", seeded); // reads back the seeded write

            var mutated = seeded.Replace("uat-run-original", "uat-run-mutated");
            Assert.Contains("uat-run-mutated", mutated); // guard: the replace changed the fetched XML
            PageWriteTools.UpdatePage(FixtureNotebook.TestVersion, mutated);

            // 5. GetPage(all) reads back step 4: mutation present AND title preserved.
            var after = PageReadTools.GetPage(FixtureNotebook.TestVersion, textPageId, "all");
            Assert.Contains("uat-run-mutated", after);        // mutation applied
            Assert.DoesNotContain("uat-run-original", after);  // old value gone
            Assert.Contains(pageTitle, after);                 // original title preserved

            // 6. ExtractPageFiles on a page seeded with a known PNG -> a decodable
            //    PNG lands on disk at an absolute path with the PNG magic bytes.
            imagePageId = PageWriteTools.CreatePage(FixtureNotebook.TestVersion, sectionId, "UAT Image " + suffix);
            PageWriteTools.UpdatePage(FixtureNotebook.TestVersion, BuildImagePageXml(imagePageId));

            var extractJson = FileExtractionTools.ExtractPageFiles(FixtureNotebook.TestVersion, imagePageId, outDir, "images");
            var files = JsonSerializer.Deserialize<List<ExtractedFile>>(extractJson, JsonOpts)!;
            var file = Assert.Single(files);
            Assert.Equal("image", file.Type);
            Assert.True(Path.IsPathRooted(file.Path), "returned path must be absolute");
            Assert.True(File.Exists(file.Path), "extracted file must exist on disk");
            Assert.Equal(PngMagic, File.ReadAllBytes(file.Path)[..4]);

            // 7. RenameNode -> the hierarchy reflects the new section name.
            var renamedName = "UAT Renamed " + suffix;
            SectionTools.RenameNode(FixtureNotebook.TestVersion, sectionId, renamedName);
            Assert.Contains(NotebookSections(), s => s.Id == sectionId && s.Name == renamedName);

            // 8. DeletePage -> the text page is gone from the section hierarchy.
            PageWriteTools.DeletePage(FixtureNotebook.TestVersion, textPageId);
            Assert.DoesNotContain(SectionPages(sectionId), p => p.Id == textPageId);
            var deletedTextPageId = textPageId;
            textPageId = null; // already deleted; skip the finally cleanup

            // 9. DeleteNode -> the section is gone from the notebook hierarchy.
            PageWriteTools.DeletePage(FixtureNotebook.TestVersion, imagePageId);
            imagePageId = null;
            SectionTools.DeleteNode(FixtureNotebook.TestVersion, sectionId);
            Assert.DoesNotContain(NotebookSections(), s => s.Id == sectionId);
            var deletedSectionId = sectionId;
            sectionId = null; // already deleted; skip the finally cleanup

            Assert.NotEqual(deletedTextPageId, deletedSectionId); // both lifecycle IDs were distinct
        }
        finally
        {
            // Best-effort teardown if an assertion aborted mid-lifecycle. Deleting
            // the section also removes any pages still under it.
            if (imagePageId is not null)
                try { PageWriteTools.DeletePage(FixtureNotebook.TestVersion, imagePageId); } catch { /* already gone */ }
            if (textPageId is not null)
                try { PageWriteTools.DeletePage(FixtureNotebook.TestVersion, textPageId); } catch { /* already gone */ }
            if (sectionId is not null)
                try { SectionTools.DeleteNode(FixtureNotebook.TestVersion, sectionId); } catch { /* already gone */ }
            if (Directory.Exists(outDir))
                try { Directory.Delete(outDir, recursive: true); } catch { /* harmless leftovers */ }
        }
    }

    /// <summary>Reads the current sections under the fixture notebook from live hierarchy XML.</summary>
    private IReadOnlyList<(string Id, string Name)> NotebookSections()
    {
        var xml = OneNoteSession.For(FixtureNotebook.TestClsid).GetHierarchy(
            _fx.NotebookId, OneNoteScope.HsSections, OneNoteXmlSchema.Xs2013);

        var root = XDocument.Parse(xml).Root;
        if (root is null) return Array.Empty<(string, string)>();

        return root.DescendantsAndSelf()
            .Where(e => e.Name.LocalName == "Section")
            .Select(e => (
                Id: e.Attribute("ID")?.Value ?? "",
                Name: e.Attribute("name")?.Value ?? ""))
            .ToList();
    }

    /// <summary>Reads the current pages under a section from live hierarchy XML.</summary>
    private static IReadOnlyList<PageMatch> SectionPages(string sectionId)
    {
        var xml = OneNoteSession.For(FixtureNotebook.TestClsid).GetHierarchy(
            sectionId, OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013);
        return HierarchyParser.ParsePages(xml);
    }

    /// <summary>Full page XML with a title and a single body text run for a page ID.</summary>
    private static string BuildPageXml(string pageId, string title, string body) =>
        $"<one:Page xmlns:one=\"{OneNs}\" ID=\"{pageId}\">" +
        $"<one:Title><one:OE><one:T><![CDATA[{title}]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren><one:OE>" +
        $"<one:T><![CDATA[{body}]]></one:T>" +
        "</one:OE></one:OEChildren></one:Outline></one:Page>";

    /// <summary>Full page XML with a title and a single embedded 1x1 PNG for a page ID.</summary>
    private static string BuildImagePageXml(string pageId) =>
        $"<one:Page xmlns:one=\"{OneNs}\" ID=\"{pageId}\">" +
        $"<one:Title><one:OE><one:T><![CDATA[UAT image]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren><one:OE>" +
        $"<one:Image format=\"png\"><one:Data>{OnePixelPngBase64}</one:Data></one:Image>" +
        "</one:OE></one:OEChildren></one:Outline></one:Page>";
}
