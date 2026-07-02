using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the page write tools (create / update /
/// delete). Each test creates its OWN scratch page under the shared fixture
/// section so it never mutates the read-fixture pages other tests assert against,
/// and deletes it on the way out. They early-return (skip) when the fixture could
/// not be built.
/// </summary>
[Collection("OneNote COM")]
public sealed class PageWriteToolsIntegrationTests
{
    private const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private readonly FixtureNotebook _fx;

    public PageWriteToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void CreatePage_WithTitle_AppearsInHierarchyWithThatTitle()
    {
        if (!_fx.Available) return;

        var title = "Scratch Create " + Guid.NewGuid().ToString("N");
        var pageId = PageWriteTools.CreatePage(_fx.SectionId, title);
        try
        {
            Assert.False(string.IsNullOrEmpty(pageId));
            Assert.Contains(SectionPages(), p => p.Id == pageId && p.Title == title);
        }
        finally
        {
            PageWriteTools.DeletePage(pageId);
        }
    }

    [Fact]
    public void UpdatePage_RoundTrip_AppliesMutationAndPreservesUntouchedContent()
    {
        if (!_fx.Available) return;

        var pageId = PageWriteTools.CreatePage(_fx.SectionId);
        try
        {
            // Seed the page with a title we keep and a body run we will mutate.
            PageWriteTools.UpdatePage(BuildPageXml(pageId, "Keep This Title", "old-body-value"));

            var before = PageReadTools.GetPage(pageId, "basic");
            Assert.Contains("old-body-value", before);

            // Mutate only the body run in the fetched XML, then write it back.
            var mutated = before.Replace("old-body-value", "new-body-value");
            Assert.Contains("new-body-value", mutated); // guard: the replace actually changed the fetched XML
            PageWriteTools.UpdatePage(mutated);

            var after = PageReadTools.GetPage(pageId, "basic");
            Assert.Contains("new-body-value", after);       // mutation applied
            Assert.DoesNotContain("old-body-value", after);  // old value gone
            Assert.Contains("Keep This Title", after);       // untouched content preserved
        }
        finally
        {
            PageWriteTools.DeletePage(pageId);
        }
    }

    [Fact]
    public void DeletePage_RemovesPageFromHierarchy()
    {
        if (!_fx.Available) return;

        var pageId = PageWriteTools.CreatePage(_fx.SectionId, "Scratch Delete " + Guid.NewGuid().ToString("N"));
        try
        {
            Assert.Contains(SectionPages(), p => p.Id == pageId); // present after create

            PageWriteTools.DeletePage(pageId);

            Assert.DoesNotContain(SectionPages(), p => p.Id == pageId); // gone after delete
        }
        finally
        {
            // Best-effort cleanup if an assertion above aborted before the delete.
            try { PageWriteTools.DeletePage(pageId); } catch { /* already deleted */ }
        }
    }

    /// <summary>Reads the current pages under the fixture section from live hierarchy XML.</summary>
    private IReadOnlyList<PageMatch> SectionPages()
    {
        var xml = OneNoteSession.Instance.GetHierarchy(
            _fx.SectionId, OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013);
        return HierarchyParser.ParsePages(xml);
    }

    /// <summary>Full page XML with a title and a single body text run for a page ID.</summary>
    private static string BuildPageXml(string pageId, string title, string body) =>
        $"<one:Page xmlns:one=\"{OneNs}\" ID=\"{pageId}\">" +
        $"<one:Title><one:OE><one:T><![CDATA[{title}]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren><one:OE>" +
        $"<one:T><![CDATA[{body}]]></one:T>" +
        "</one:OE></one:OEChildren></one:Outline></one:Page>";
}
