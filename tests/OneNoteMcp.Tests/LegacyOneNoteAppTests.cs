using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fakes;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free tests for <see cref="LegacyOneNoteApp"/> (OneNote 2007 / v12). Proves:
/// (1) supported calls map onto v12's SHORTER signatures — the XMLSchema argument
/// and the modern-only trailing booleans are dropped, never forwarded; and
/// (2) the six IApplication2–4 methods absent on v12 throw
/// <see cref="NotSupportedInVersionException"/> for version 12 WITHOUT ever
/// touching the RCW (the fake's <c>LastMethod</c> stays null). A recording fake
/// v12 IApplication stands in for the RCW.
/// </summary>
public sealed class LegacyOneNoteAppTests
{
    private static (LegacyOneNoteApp app, FakeLegacyOneNoteApp fake) NewApp(string ret = "")
    {
        var fake = new FakeLegacyOneNoteApp { ReturnValue = ret };
        return (new LegacyOneNoteApp(fake), fake);
    }

    [Fact]
    public void GetHierarchy_DropsSchema_ReturnsXml()
    {
        var (app, fake) = NewApp("<H/>");
        var xml = app.GetHierarchy("node-1", OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013);
        Assert.Equal("<H/>", xml);
        Assert.Equal("GetHierarchy", fake.LastMethod);
        // Schema is NOT forwarded — only start + scope reach the v12 RCW.
        Assert.Equal(new object?[] { "node-1", OneNoteScope.HsPages }, fake.LastArgs);
    }

    [Fact]
    public void FindMeta_DropsSchema()
    {
        var (app, fake) = NewApp("<M/>");
        var xml = app.FindMeta("meta-name", OneNoteXmlSchema.Xs2013);
        Assert.Equal("<M/>", xml);
        Assert.Equal("FindMeta", fake.LastMethod);
        // start(default ""), search, fUnindexed(false) — no schema.
        Assert.Equal(new object?[] { "", "meta-name", false }, fake.LastArgs);
    }

    [Fact]
    public void UpdatePageContent_DropsSchemaAndForce()
    {
        var (app, fake) = NewApp();
        app.UpdatePageContent("<changes/>", OneNoteXmlSchema.Xs2013);
        Assert.Equal("UpdatePageContent", fake.LastMethod);
        // v12 UpdatePageContent takes only (xml, date) — 2 args, no schema/force.
        Assert.Equal(2, fake.LastArgs.Count);
        Assert.Equal("<changes/>", fake.LastArgs[0]);
    }

    [Fact]
    public void DeleteHierarchy_DropsForce()
    {
        var (app, fake) = NewApp();
        app.DeleteHierarchy("obj-1");
        Assert.Equal("DeleteHierarchy", fake.LastMethod);
        // v12 DeleteHierarchy takes (objectId, date) — no deletePermanently bool.
        Assert.Equal(2, fake.LastArgs.Count);
        Assert.Equal("obj-1", fake.LastArgs[0]);
    }

    [Fact]
    public void DeletePageContent_DropsForce()
    {
        var (app, fake) = NewApp();
        app.DeletePageContent("page-1", "obj-2");
        Assert.Equal("DeletePageContent", fake.LastMethod);
        // v12 DeletePageContent takes (pageId, objectId, date) — no force bool.
        Assert.Equal(3, fake.LastArgs.Count);
        Assert.Equal("page-1", fake.LastArgs[0]);
        Assert.Equal("obj-2", fake.LastArgs[1]);
    }

    [Fact]
    public void GetHierarchyParent_Delegates_ReturnsParent()
    {
        var (app, fake) = NewApp("parent-9");
        var parent = app.GetHierarchyParent("child-1");
        Assert.Equal("parent-9", parent);
        Assert.Equal("GetHierarchyParent", fake.LastMethod);
        Assert.Equal(new object?[] { "child-1" }, fake.LastArgs);
    }

    [Fact]
    public void CreateNewPage_Delegates_ReturnsId()
    {
        var (app, fake) = NewApp("new-page");
        var id = app.CreateNewPage("section-1", OneNoteNewPageStyle.NpsBlankPageWithTitle);
        Assert.Equal("new-page", id);
        Assert.Equal("CreateNewPage", fake.LastMethod);
        Assert.Equal(new object?[] { "section-1", OneNoteNewPageStyle.NpsBlankPageWithTitle }, fake.LastArgs);
    }

    // ── The six methods absent on v12 must throw without touching the RCW ──────

    private static void AssertNotSupported(FakeLegacyOneNoteApp fake, string method, Action call)
    {
        var ex = Assert.Throws<NotSupportedInVersionException>(call);
        Assert.Equal(12, ex.VersionMajor);
        Assert.Equal(method, ex.MethodName);
        Assert.Null(fake.LastMethod); // never dispatched to COM
    }

    [Fact]
    public void NavigateToUrl_Throws()
    {
        var (app, fake) = NewApp();
        AssertNotSupported(fake, nameof(app.NavigateToUrl), () => app.NavigateToUrl("onenote:x"));
    }

    [Fact]
    public void GetWebHyperlinkToObject_Throws()
    {
        var (app, fake) = NewApp();
        AssertNotSupported(fake, nameof(app.GetWebHyperlinkToObject),
            () => app.GetWebHyperlinkToObject("hier-1", "obj-2"));
    }

    [Fact]
    public void MergeFiles_Throws()
    {
        var (app, fake) = NewApp();
        AssertNotSupported(fake, nameof(app.MergeFiles),
            () => app.MergeFiles("base", "client", "server", "target"));
    }

    [Fact]
    public void MergeSections_Throws()
    {
        var (app, fake) = NewApp();
        AssertNotSupported(fake, nameof(app.MergeSections), () => app.MergeSections("src", "dst"));
    }

    [Fact]
    public void SyncHierarchy_Throws()
    {
        var (app, fake) = NewApp();
        AssertNotSupported(fake, nameof(app.SyncHierarchy), () => app.SyncHierarchy("hier-9"));
    }

    [Fact]
    public void SetFilingLocation_Throws()
    {
        var (app, fake) = NewApp();
        AssertNotSupported(fake, nameof(app.SetFilingLocation),
            () => app.SetFilingLocation(0, 0, "section-1"));
    }
}
