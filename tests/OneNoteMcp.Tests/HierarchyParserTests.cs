using System.Linq;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure-logic tests for <see cref="HierarchyParser"/>. No COM: these exercise the
/// XML parsing against clean-room hierarchy samples in the OneNote 2013 schema,
/// so they run on every machine regardless of whether OneNote is installed.
/// </summary>
public sealed class HierarchyParserTests
{
    private const string Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private const string NotebooksXml =
        "<?xml version=\"1.0\"?>" +
        "<one:Notebooks xmlns:one=\"" + Ns + "\">" +
        "<one:Notebook name=\"Personal\" ID=\"{N1}\" path=\"C:\\nb\\Personal\\\" " +
            "lastModifiedTime=\"2026-06-01T10:00:00.000Z\" color=\"#FFFFFF\" />" +
        "<one:Notebook name=\"Work\" ID=\"{N2}\" path=\"C:\\nb\\Work\\\" " +
            "lastModifiedTime=\"2026-06-02T12:00:00.000Z\" readOnly=\"true\" isUnread=\"true\" />" +
        "</one:Notebooks>";

    private const string PagesXml =
        "<?xml version=\"1.0\"?>" +
        "<one:Notebooks xmlns:one=\"" + Ns + "\">" +
        "<one:Notebook name=\"Personal\" ID=\"{N1}\" path=\"C:\\nb\\Personal\\\">" +
        "<one:Section name=\"Sec1\" ID=\"{S1}\" path=\"C:\\nb\\Personal\\Sec1.one\">" +
        "<one:Page name=\"Page A\" ID=\"{P1}\" pageLevel=\"1\" />" +
        "<one:Page name=\"Page B\" ID=\"{P2}\" pageLevel=\"1\" />" +
        "</one:Section>" +
        "<one:SectionGroup name=\"Grp\" ID=\"{G1}\">" +
        "<one:Section name=\"Sec2\" ID=\"{S2}\">" +
        "<one:Page name=\"Page C\" ID=\"{P3}\" />" +
        "</one:Section>" +
        "</one:SectionGroup>" +
        "</one:Notebook>" +
        "</one:Notebooks>";

    [Fact]
    public void ParseNotebooks_ReturnsIdNamePathForEachNotebook()
    {
        var notebooks = HierarchyParser.ParseNotebooks(NotebooksXml);

        Assert.Equal(2, notebooks.Count);

        var personal = notebooks.Single(n => n.Name == "Personal");
        Assert.Equal("{N1}", personal.Id);
        Assert.Equal("C:\\nb\\Personal\\", personal.Path);
        Assert.Equal("2026-06-01T10:00:00.000Z", personal.LastModified);
    }

    [Fact]
    public void ParseNotebooks_MapsReadOnlyAndUnreadFlags()
    {
        var notebooks = HierarchyParser.ParseNotebooks(NotebooksXml);

        var personal = notebooks.Single(n => n.Name == "Personal");
        Assert.False(personal.IsReadOnly);
        Assert.False(personal.IsUnread);

        var work = notebooks.Single(n => n.Name == "Work");
        Assert.True(work.IsReadOnly);
        Assert.True(work.IsUnread);
    }

    [Fact]
    public void ParseNotebooks_EmptyWhenNoNotebooks()
    {
        var xml = "<one:Notebooks xmlns:one=\"" + Ns + "\" />";
        Assert.Empty(HierarchyParser.ParseNotebooks(xml));
    }

    [Fact]
    public void ParsePages_ReturnsEachPageWithItsSectionId()
    {
        var pages = HierarchyParser.ParsePages(PagesXml);

        Assert.Equal(3, pages.Count);

        var a = pages.Single(p => p.Id == "{P1}");
        Assert.Equal("Page A", a.Title);
        Assert.Equal("{S1}", a.SectionId);

        // Page nested inside a SectionGroup still resolves to its immediate section.
        var c = pages.Single(p => p.Id == "{P3}");
        Assert.Equal("Page C", c.Title);
        Assert.Equal("{S2}", c.SectionId);
    }

    [Fact]
    public void ParsePages_EmptyWhenNoPages()
    {
        Assert.Empty(HierarchyParser.ParsePages(NotebooksXml));
    }
}
