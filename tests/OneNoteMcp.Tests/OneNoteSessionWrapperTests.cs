using System.Runtime.InteropServices;
using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for the OneNoteSession pass-through wrappers added for full
/// IApplication coverage. Each test injects a fake COM app (a duck-typed POCO
/// dispatched via the late-bound InvokeFake path) that records the arguments it
/// received and writes back canned out-params, proving the wrapper marshals the
/// right positional arguments and returns the right out-value — no real OneNote.
/// </summary>
[Collection("OneNote COM")]
public sealed class OneNoteSessionWrapperTests
{
    /// <summary>
    /// One fake standing in for the whole IApplication surface these wrappers use.
    /// Out-params are declared <c>ref string</c> so the late-bound InvokeMember path
    /// writes the canned return back into the caller's argument array. Every call
    /// records its inputs for assertion.
    /// </summary>
    private sealed class RecordingApp
    {
        public List<object?> LastArgs = new();
        public string? LastMethod;
        public string ReturnValue = "";

        private void Record(string method, params object?[] inputs)
        {
            LastMethod = method;
            LastArgs = inputs.ToList();
        }

        public void GetBinaryPageContent(string pageId, string callbackId, ref string b64)
        {
            Record(nameof(GetBinaryPageContent), pageId, callbackId);
            b64 = ReturnValue;
        }

        public void DeletePageContent(string pageId, string objectId, DateTime date, bool force)
            => Record(nameof(DeletePageContent), pageId, objectId, date, force);

        public void GetHierarchyParent(string objectId, ref string parentId)
        {
            Record(nameof(GetHierarchyParent), objectId);
            parentId = ReturnValue;
        }

        public void GetSpecialLocation(int slToGet, ref string path)
        {
            Record(nameof(GetSpecialLocation), slToGet);
            path = ReturnValue;
        }

        public void NavigateTo(string hierarchyObjectId, string objectId, bool newWindow)
            => Record(nameof(NavigateTo), hierarchyObjectId, objectId, newWindow);

        public void NavigateToUrl(string url, bool newWindow)
            => Record(nameof(NavigateToUrl), url, newWindow);

        public void GetHyperlinkToObject(string hierarchyId, string pageContentObjectId, ref string link)
        {
            Record(nameof(GetHyperlinkToObject), hierarchyId, pageContentObjectId);
            link = ReturnValue;
        }

        public void GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId, ref string link)
        {
            Record(nameof(GetWebHyperlinkToObject), hierarchyId, pageContentObjectId);
            link = ReturnValue;
        }

        public void FindMeta(string startNodeId, string searchName, ref string xml,
            bool includeUnindexed, int xsSchema)
        {
            Record(nameof(FindMeta), startNodeId, searchName, includeUnindexed, xsSchema);
            xml = ReturnValue;
        }

        public void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile)
            => Record(nameof(MergeFiles), baseFile, clientFile, serverFile, targetFile);

        public void MergeSections(string sourceId, string destId)
            => Record(nameof(MergeSections), sourceId, destId);

        public void SyncHierarchy(string hierarchyId)
            => Record(nameof(SyncHierarchy), hierarchyId);

        public void SetFilingLocation(int flToSet, int fltToSet, string sectionId)
            => Record(nameof(SetFilingLocation), flToSet, fltToSet, sectionId);
    }

    /// <summary>Fake whose GetBinaryPageContent always reports a dead COM proxy.</summary>
    private sealed class DeadBinaryApp
    {
        public void GetBinaryPageContent(string pageId, string callbackId, ref string b64)
            => throw new COMException("Server unavailable", unchecked((int)0x800706BA));
    }

    private static OneNoteSession SessionWith(RecordingApp app) => new(() => app);

    [Fact]
    public void GetBinaryPageContent_ReturnsBase64FromApp()
    {
        var app = new RecordingApp { ReturnValue = "QUJD" };
        using var session = SessionWith(app);

        var result = session.GetBinaryPageContent("page-1", "cb-9");

        Assert.Equal("QUJD", result);
        Assert.Equal("GetBinaryPageContent", app.LastMethod);
        Assert.Equal(new object?[] { "page-1", "cb-9" }, app.LastArgs);
    }

    [Fact]
    public void DeletePageContent_PassesPageAndObjectId()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.DeletePageContent("page-1", "obj-2");

        Assert.Equal("DeletePageContent", app.LastMethod);
        Assert.Equal("page-1", app.LastArgs[0]);
        Assert.Equal("obj-2", app.LastArgs[1]);
    }

    [Fact]
    public void GetHierarchyParent_ReturnsParentId()
    {
        var app = new RecordingApp { ReturnValue = "parent-42" };
        using var session = SessionWith(app);

        var result = session.GetHierarchyParent("child-1");

        Assert.Equal("parent-42", result);
        Assert.Equal(new object?[] { "child-1" }, app.LastArgs);
    }

    [Fact]
    public void GetSpecialLocation_ReturnsPath()
    {
        var app = new RecordingApp { ReturnValue = @"C:\Notebooks" };
        using var session = SessionWith(app);

        var result = session.GetSpecialLocation(OneNoteSpecialLocation.SlDefaultNotebookFolder);

        Assert.Equal(@"C:\Notebooks", result);
        Assert.Equal(OneNoteSpecialLocation.SlDefaultNotebookFolder, app.LastArgs[0]);
    }

    [Fact]
    public void NavigateTo_PassesIds()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.NavigateTo("hier-1", "obj-2", newWindow: true);

        Assert.Equal("NavigateTo", app.LastMethod);
        Assert.Equal(new object?[] { "hier-1", "obj-2", true }, app.LastArgs);
    }

    [Fact]
    public void NavigateToUrl_PassesUrl()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.NavigateToUrl("https://example.com");

        Assert.Equal("NavigateToUrl", app.LastMethod);
        Assert.Equal(new object?[] { "https://example.com", false }, app.LastArgs);
    }

    [Fact]
    public void GetHyperlinkToObject_ReturnsLink()
    {
        var app = new RecordingApp { ReturnValue = "onenote:link" };
        using var session = SessionWith(app);

        var result = session.GetHyperlinkToObject("hier-1", "obj-2");

        Assert.Equal("onenote:link", result);
        Assert.Equal(new object?[] { "hier-1", "obj-2" }, app.LastArgs);
    }

    [Fact]
    public void GetWebHyperlinkToObject_ReturnsLink()
    {
        var app = new RecordingApp { ReturnValue = "https://onenote/link" };
        using var session = SessionWith(app);

        var result = session.GetWebHyperlinkToObject("hier-1", "obj-2");

        Assert.Equal("https://onenote/link", result);
        Assert.Equal(new object?[] { "hier-1", "obj-2" }, app.LastArgs);
    }

    [Fact]
    public void FindMeta_ReturnsXml()
    {
        var app = new RecordingApp { ReturnValue = "<Hierarchy />" };
        using var session = SessionWith(app);

        var result = session.FindMeta("meta-name", OneNoteXmlSchema.Xs2013);

        Assert.Equal("<Hierarchy />", result);
        Assert.Equal("FindMeta", app.LastMethod);
        // startNodeId defaults empty, searchName passed through.
        Assert.Equal("", app.LastArgs[0]);
        Assert.Equal("meta-name", app.LastArgs[1]);
    }

    [Fact]
    public void MergeFiles_PassesFourPaths()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.MergeFiles("base", "client", "server", "target");

        Assert.Equal("MergeFiles", app.LastMethod);
        Assert.Equal(new object?[] { "base", "client", "server", "target" }, app.LastArgs);
    }

    [Fact]
    public void MergeSections_PassesSourceAndDest()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.MergeSections("src-1", "dst-2");

        Assert.Equal("MergeSections", app.LastMethod);
        Assert.Equal(new object?[] { "src-1", "dst-2" }, app.LastArgs);
    }

    [Fact]
    public void SyncHierarchy_PassesId()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.SyncHierarchy("hier-9");

        Assert.Equal("SyncHierarchy", app.LastMethod);
        Assert.Equal(new object?[] { "hier-9" }, app.LastArgs);
    }

    [Fact]
    public void SetFilingLocation_PassesEnumsAndSection()
    {
        var app = new RecordingApp();
        using var session = SessionWith(app);

        session.SetFilingLocation(
            OneNoteFilingLocation.FlEMail,
            OneNoteFilingLocationType.FltNamedSectionNewPage,
            "section-1");

        Assert.Equal("SetFilingLocation", app.LastMethod);
        Assert.Equal(new object?[]
        {
            OneNoteFilingLocation.FlEMail,
            OneNoteFilingLocationType.FltNamedSectionNewPage,
            "section-1",
        }, app.LastArgs);
    }

    [Fact]
    public void GetBinaryPageContent_DeadProxy_Recreates()
    {
        var factoryCallCount = 0;
        using var session = new OneNoteSession(() =>
        {
            factoryCallCount++;
            return factoryCallCount == 1 ? new DeadBinaryApp() : (object)new RecordingApp { ReturnValue = "OK" };
        });

        // First call: dead proxy throws 0x800706BA and is invalidated.
        Assert.Throws<COMException>(() => session.GetBinaryPageContent("p", "cb"));
        Assert.Equal(1, factoryCallCount);

        // Second call: factory recreates a live app — recreate-on-death confirmed
        // for the new wrapper, which rides the shared Dispatch path.
        var result = session.GetBinaryPageContent("p", "cb");
        Assert.Equal("OK", result);
        Assert.Equal(2, factoryCallCount);
    }
}
