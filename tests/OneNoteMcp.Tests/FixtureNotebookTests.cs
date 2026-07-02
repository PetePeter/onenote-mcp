using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fixtures;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Self-tests proving the shared <see cref="FixtureNotebook"/> builds real,
/// discoverable content and tears itself down cleanly. Every test skips (early
/// return) when the fixture is unavailable so machines without OneNote pass.
/// </summary>
[Collection("OneNote COM")]
public sealed class FixtureNotebookTests
{
    private readonly FixtureNotebook _fx;

    public FixtureNotebookTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void Fixture_NotebookAppearsInHierarchy()
    {
        if (!_fx.Available) return;

        var xml = OneNoteSession.Instance.GetHierarchy(
            "", OneNoteScope.HsSections, OneNoteXmlSchema.Xs2013);

        Assert.Contains(_fx.NotebookId, xml);
    }

    [Fact]
    public void Fixture_TextPage_HasKnownTitleAndText()
    {
        if (!_fx.Available) return;

        var content = OneNoteSession.Instance.GetPageContent(
            _fx.TextPageId, 0, OneNoteXmlSchema.Xs2013);

        Assert.Contains(_fx.KnownTitle, content);
        Assert.Contains(_fx.KnownText, content);
    }

    [Fact]
    public void Fixture_ImagePage_ContainsImage()
    {
        if (!_fx.Available) return;

        var content = OneNoteSession.Instance.GetPageContent(
            _fx.ImagePageId, 0, OneNoteXmlSchema.Xs2013);

        Assert.Contains("one:Image", content);
    }

    [Fact]
    public void Fixture_Dispose_RemovesNotebookAndFiles()
    {
        // Self-contained lifecycle: builds and disposes its OWN fixture so the
        // shared one is untouched.
        if (!OneNoteSession.IsComAvailable) return;

        var f = new FixtureNotebook();
        if (!f.Available) return;

        var dir = f.Directory;
        var id = f.NotebookId;

        f.Dispose();

        Assert.False(System.IO.Directory.Exists(dir));

        var hierarchy = OneNoteSession.Instance.GetHierarchy(
            "", OneNoteScope.HsSections, OneNoteXmlSchema.Xs2013);
        Assert.DoesNotContain(id, hierarchy);
    }
}
