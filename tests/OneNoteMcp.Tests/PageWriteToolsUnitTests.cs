using System.Runtime.InteropServices;
using System.Xml.Linq;
using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fakes;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Regression unit tests (no real COM) for two confirmed page-write bugs:
///  1. CreateNewPage duplicated pages because the non-idempotent create was retried
///     on a transient COM error that fired AFTER the page was already created.
///  2. The create-with-title path emitted the 2013 namespace/schema unconditionally,
///     so OneNote 2007 (v12) rejected the title write (0x80042001) and orphaned a page.
/// Uses a recording fake IOneNoteApp — fakes over mocks, real production classes.
/// </summary>
[Collection("OneNote COM")]
public sealed class PageWriteToolsUnitTests : IDisposable
{
    // Canonical CLSIDs from the version catalog.
    private const string Clsid2016 = "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}";

    private const int RpcECallRejected = unchecked((int)0x80010001);
    private const string Ns2007 = "http://schemas.microsoft.com/office/onenote/2007/onenote";
    private const string Ns2013 = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    public PageWriteToolsUnitTests() => OneNoteSession.ResetForTests();
    public void Dispose() => OneNoteSession.ResetForTests();

    /// <summary>
    /// Recording fake whose CreateNewPage throws a transient COM error on the FIRST
    /// call and would succeed on any later call. Also counts CreateNewPage +
    /// DeleteHierarchy invocations so tests can prove call multiplicity.
    /// </summary>
    private sealed class SequencingFake : IOneNoteApp
    {
        public int CreateCalls { get; private set; }
        public int DeleteCalls { get; private set; }
        public bool TransientOnFirstCreate { get; set; }
        public bool FailTitleWrite { get; set; }

        public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault)
        {
            CreateCalls++;
            if (TransientOnFirstCreate && CreateCalls == 1)
                throw new COMException("busy", RpcECallRejected);
            return "page-id";
        }

        public void UpdatePageContent(string pageChangesXml, int xmlSchema)
        {
            if (FailTitleWrite)
                throw new COMException("bad schema", unchecked((int)0x80042001));
        }

        public void DeleteHierarchy(string objectId) => DeleteCalls++;

        // ── Unused surface ───────────────────────────────────────────────────────
        public string GetHierarchy(string s, int a, int b) => "";
        public string FindPages(string s, int a, string b = "") => "";
        public string GetPageContent(string a, int b, int c) => "";
        public string OpenHierarchy(string a, string b, int c) => "";
        public void Publish(string a, string b, int c, string d = "") { }
        public void CloseNotebook(string a) { }
        public void UpdateHierarchy(string a, int b) { }
        public string GetBinaryPageContent(string a, string b) => "";
        public void DeletePageContent(string a, string b) { }
        public string GetHierarchyParent(string a) => "";
        public string GetSpecialLocation(int a) => "";
        public void NavigateTo(string a, string b = "", bool c = false) { }
        public string GetHyperlinkToObject(string a, string b = "") => "";
        public string FindMeta(string a, int b, string c = "") => "";
        public void NavigateToUrl(string a, bool b = false) { }
        public string GetWebHyperlinkToObject(string a, string b = "") => "";
        public void MergeFiles(string a, string b, string c, string d) { }
        public void MergeSections(string a, string b) { }
        public void SyncHierarchy(string a) { }
        public void SetFilingLocation(int a, int b, string c) { }
    }

    // ── BUG 1: no retry on non-idempotent create ─────────────────────────────────

    [Fact]
    public void CreateNewPage_TransientComError_IsNotRetried_SoPageIsNotDuplicated()
    {
        var fake = new SequencingFake { TransientOnFirstCreate = true };
        using var session = new OneNoteSession(() => fake);

        // A transient error on create must propagate after exactly ONE attempt.
        // Retrying (the old behaviour) would call CreateNewPage up to 5 times and,
        // since OneNote may have already created the page, duplicate it.
        Assert.Throws<COMException>(() => session.CreateNewPage("section-id"));
        Assert.Equal(1, fake.CreateCalls);
    }

    [Fact]
    public void GetHierarchy_TransientComError_IsStillRetried()
    {
        // Guard: the no-retry change must be scoped to non-idempotent creates only;
        // idempotent reads must keep retrying on transient errors.
        int calls = 0;
        var fake = new RetryReadFake(() =>
        {
            calls++;
            if (calls == 1) throw new COMException("busy", RpcECallRejected);
            return "<Notebooks/>";
        });
        using var session = new OneNoteSession(() => fake);

        var xml = session.GetHierarchy("", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013);

        Assert.Equal("<Notebooks/>", xml);
        Assert.Equal(2, calls); // retried once, then succeeded
    }

    private sealed class RetryReadFake : IOneNoteApp
    {
        private readonly Func<string> _getHierarchy;
        public RetryReadFake(Func<string> getHierarchy) => _getHierarchy = getHierarchy;
        public string GetHierarchy(string s, int a, int b) => _getHierarchy();
        public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault) => "";
        public string FindPages(string s, int a, string b = "") => "";
        public string GetPageContent(string a, int b, int c) => "";
        public void UpdatePageContent(string a, int b) { }
        public string OpenHierarchy(string a, string b, int c) => "";
        public void Publish(string a, string b, int c, string d = "") { }
        public void CloseNotebook(string a) { }
        public void DeleteHierarchy(string a) { }
        public void UpdateHierarchy(string a, int b) { }
        public string GetBinaryPageContent(string a, string b) => "";
        public void DeletePageContent(string a, string b) { }
        public string GetHierarchyParent(string a) => "";
        public string GetSpecialLocation(int a) => "";
        public void NavigateTo(string a, string b = "", bool c = false) { }
        public string GetHyperlinkToObject(string a, string b = "") => "";
        public string FindMeta(string a, int b, string c = "") => "";
        public void NavigateToUrl(string a, bool b = false) { }
        public string GetWebHyperlinkToObject(string a, string b = "") => "";
        public void MergeFiles(string a, string b, string c, string d) { }
        public void MergeSections(string a, string b) { }
        public void SyncHierarchy(string a) { }
        public void SetFilingLocation(int a, int b, string c) { }
    }

    // ── BUG 2: version-correct title namespace/schema ────────────────────────────

    [Fact]
    public void BuildTitleXml_Major12_Uses2007Namespace()
    {
        var xml = PageWriteTools.BuildTitleXml("page-id", "Hello", 12);
        var root = XElement.Parse(xml);

        Assert.Equal(Ns2007, root.Name.NamespaceName);
        Assert.Equal("Hello", root.Descendants(XName.Get("T", Ns2007)).Single().Value);
    }

    [Fact]
    public void BuildTitleXml_Major16_Uses2013Namespace()
    {
        var xml = PageWriteTools.BuildTitleXml("page-id", "Hello", 16);
        var root = XElement.Parse(xml);

        Assert.Equal(Ns2013, root.Name.NamespaceName);
    }

    [Fact]
    public void CreatePage_TitleWriteFails_DeletesOrphanPage()
    {
        // If the title write fails, the just-created page must be cleaned up so no
        // orphan "Untitled" page is left behind.
        var fake = new SequencingFake { FailTitleWrite = true };
        OneNoteSession.AppFactoryOverride = _ => fake;

        var result = PageWriteTools.CreatePage("2016", "section-id", "A Title");

        // Guard wraps the failure into an error string, and the orphan is deleted.
        Assert.Equal(1, fake.CreateCalls);
        Assert.Equal(1, fake.DeleteCalls);
    }
}
