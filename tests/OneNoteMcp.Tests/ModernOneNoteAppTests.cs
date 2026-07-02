using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fakes;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free tests for <see cref="ModernOneNoteApp"/>: it must delegate every
/// <see cref="IOneNoteApp"/> call to the matching modern IApplication method with
/// the right positional arguments (including the XMLSchema pass-through the modern
/// surface keeps), and — unlike the legacy adapter — the six IApplication2–4
/// methods must delegate rather than throw. A recording fake IApplication stands
/// in for the RCW.
/// </summary>
public sealed class ModernOneNoteAppTests
{
    private static (ModernOneNoteApp app, FakeModernOneNoteApp fake) NewApp(string ret = "")
    {
        var fake = new FakeModernOneNoteApp { ReturnValue = ret };
        return (new ModernOneNoteApp(fake), fake);
    }

    [Fact]
    public void GetHierarchy_DelegatesWithScopeAndSchema_ReturnsXml()
    {
        var (app, fake) = NewApp("<H/>");
        var xml = app.GetHierarchy("node-1", OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013);
        Assert.Equal("<H/>", xml);
        Assert.Equal("GetHierarchy", fake.LastMethod);
        Assert.Equal(new object?[] { "node-1", OneNoteScope.HsPages, OneNoteXmlSchema.Xs2013 }, fake.LastArgs);
    }

    [Fact]
    public void FindPages_DelegatesWithSchema_ReturnsXml()
    {
        var (app, fake) = NewApp("<P/>");
        var xml = app.FindPages("query", OneNoteXmlSchema.Xs2010, "start-1");
        Assert.Equal("<P/>", xml);
        Assert.Equal("FindPages", fake.LastMethod);
        // start, search, fUnindexed(false), fDisplay(false), schema
        Assert.Equal(new object?[] { "start-1", "query", false, false, OneNoteXmlSchema.Xs2010 }, fake.LastArgs);
    }

    [Fact]
    public void GetPageContent_DelegatesWithPageInfoAndSchema()
    {
        var (app, fake) = NewApp("<page/>");
        var xml = app.GetPageContent("page-1", 3, OneNoteXmlSchema.Xs2013);
        Assert.Equal("<page/>", xml);
        Assert.Equal("GetPageContent", fake.LastMethod);
        Assert.Equal(new object?[] { "page-1", 3, OneNoteXmlSchema.Xs2013 }, fake.LastArgs);
    }

    [Fact]
    public void UpdatePageContent_DelegatesWithSchema()
    {
        var (app, fake) = NewApp();
        app.UpdatePageContent("<changes/>", OneNoteXmlSchema.Xs2013);
        Assert.Equal("UpdatePageContent", fake.LastMethod);
        Assert.Equal("<changes/>", fake.LastArgs[0]);
        Assert.Contains(OneNoteXmlSchema.Xs2013, fake.LastArgs.Cast<object>());
    }

    [Fact]
    public void CreateNewPage_DefaultStyle_ReturnsId()
    {
        var (app, fake) = NewApp("new-page");
        var id = app.CreateNewPage("section-1");
        Assert.Equal("new-page", id);
        Assert.Equal("CreateNewPage", fake.LastMethod);
        Assert.Equal(new object?[] { "section-1", OneNoteNewPageStyle.NpsDefault }, fake.LastArgs);
    }

    [Fact]
    public void OpenHierarchy_ReturnsObjectId()
    {
        var (app, fake) = NewApp("obj-42");
        var id = app.OpenHierarchy("path", "rel", OneNoteCreateFileType.CftSection);
        Assert.Equal("obj-42", id);
        Assert.Equal("OpenHierarchy", fake.LastMethod);
        Assert.Equal(new object?[] { "path", "rel", OneNoteCreateFileType.CftSection }, fake.LastArgs);
    }

    [Fact]
    public void GetBinaryPageContent_ReturnsBase64()
    {
        var (app, fake) = NewApp("QUJD");
        var b64 = app.GetBinaryPageContent("page-1", "cb-1");
        Assert.Equal("QUJD", b64);
        Assert.Equal("GetBinaryPageContent", fake.LastMethod);
        Assert.Equal(new object?[] { "page-1", "cb-1" }, fake.LastArgs);
    }

    [Fact]
    public void CloseNotebook_DoesNotForce()
    {
        var (app, fake) = NewApp();
        app.CloseNotebook("nb-1");
        Assert.Equal("CloseNotebook", fake.LastMethod);
        // Must mirror OneNoteSession, which omits force (COM default false) — a
        // forced close would tear down a notebook with unsaved state.
        Assert.Equal(new object?[] { "nb-1", false }, fake.LastArgs);
    }

    [Fact]
    public void GetSpecialLocation_ReturnsPath()
    {
        var (app, fake) = NewApp(@"C:\Backups");
        var path = app.GetSpecialLocation(OneNoteSpecialLocation.SlBackUpFolder);
        Assert.Equal(@"C:\Backups", path);
        Assert.Equal(new object?[] { OneNoteSpecialLocation.SlBackUpFolder }, fake.LastArgs);
    }

    // ── The six modern-only methods must DELEGATE (never throw) on the modern adapter ──

    [Fact]
    public void NavigateToUrl_Delegates()
    {
        var (app, fake) = NewApp();
        app.NavigateToUrl("onenote:x", newWindow: true);
        Assert.Equal("NavigateToUrl", fake.LastMethod);
        Assert.Equal(new object?[] { "onenote:x", true }, fake.LastArgs);
    }

    [Fact]
    public void GetWebHyperlinkToObject_Delegates_ReturnsLink()
    {
        var (app, fake) = NewApp("https://onenote/x");
        var link = app.GetWebHyperlinkToObject("hier-1", "obj-2");
        Assert.Equal("https://onenote/x", link);
        Assert.Equal("GetWebHyperlinkToObject", fake.LastMethod);
        Assert.Equal(new object?[] { "hier-1", "obj-2" }, fake.LastArgs);
    }

    [Fact]
    public void MergeFiles_Delegates()
    {
        var (app, fake) = NewApp();
        app.MergeFiles("base", "client", "server", "target");
        Assert.Equal("MergeFiles", fake.LastMethod);
        Assert.Equal(new object?[] { "base", "client", "server", "target" }, fake.LastArgs);
    }

    [Fact]
    public void MergeSections_Delegates()
    {
        var (app, fake) = NewApp();
        app.MergeSections("src", "dst");
        Assert.Equal("MergeSections", fake.LastMethod);
        Assert.Equal(new object?[] { "src", "dst" }, fake.LastArgs);
    }

    [Fact]
    public void SyncHierarchy_Delegates()
    {
        var (app, fake) = NewApp();
        app.SyncHierarchy("hier-9");
        Assert.Equal("SyncHierarchy", fake.LastMethod);
        Assert.Equal(new object?[] { "hier-9" }, fake.LastArgs);
    }

    [Fact]
    public void SetFilingLocation_Delegates()
    {
        var (app, fake) = NewApp();
        app.SetFilingLocation(OneNoteFilingLocation.FlEMail,
            OneNoteFilingLocationType.FltCurrentPage, "section-1");
        Assert.Equal("SetFilingLocation", fake.LastMethod);
        Assert.Equal(new object?[]
        {
            OneNoteFilingLocation.FlEMail,
            OneNoteFilingLocationType.FltCurrentPage,
            "section-1",
        }, fake.LastArgs);
    }
}
