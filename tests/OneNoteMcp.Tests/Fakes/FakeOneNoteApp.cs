using System.Runtime.InteropServices;
using OneNoteMcp.Interop;

namespace OneNoteMcp.Tests.Fakes;

/// <summary>
/// Test double for IOneNoteApp. Records which method was last called and its args.
/// Configure ThrowIfCalled=true to guard that COM is never reached (gate tests).
/// Configure DeadProxy=true to simulate 0x800706BA crashes for recreate tests.
/// </summary>
public sealed class FakeOneNoteApp : IOneNoteApp
{
    /// <summary>Name of the last method invoked, or null if none yet.</summary>
    public string? LastMethod { get; private set; }

    /// <summary>Args passed to the last method invoked.</summary>
    public List<object?> LastArgs { get; private set; } = new();

    /// <summary>Value returned by all string-returning methods.</summary>
    public string ReturnValue { get; set; } = "";

    /// <summary>When true, every method call throws InvalidOperationException.</summary>
    public bool ThrowIfCalled { get; set; }

    /// <summary>When true, every method call throws the "server unavailable" COMException.</summary>
    public bool DeadProxy { get; set; }

    private void Record(string method, params object?[] args)
    {
        LastMethod = method;
        LastArgs = args.ToList();
    }

    private void Guard()
    {
        if (ThrowIfCalled) throw new InvalidOperationException("COM must not be reached");
        if (DeadProxy) throw new COMException("Server unavailable", unchecked((int)0x800706BA));
    }

    private string GuardReturn(string method, params object?[] args)
    {
        Guard();
        Record(method, args);
        return ReturnValue;
    }

    private void GuardVoid(string method, params object?[] args)
    {
        Guard();
        Record(method, args);
    }

    // ── IOneNoteApp implementation ───────────────────────────────────────────────

    public string GetHierarchy(string startNodeId, int scope, int xmlSchema)
        => GuardReturn(nameof(GetHierarchy), startNodeId, scope, xmlSchema);

    public string FindPages(string searchString, int xmlSchema, string startNodeId = "")
        => GuardReturn(nameof(FindPages), searchString, xmlSchema, startNodeId);

    public string GetPageContent(string pageId, int pageInfo, int xmlSchema)
        => GuardReturn(nameof(GetPageContent), pageId, pageInfo, xmlSchema);

    public void UpdatePageContent(string pageChangesXml, int xmlSchema)
        => GuardVoid(nameof(UpdatePageContent), pageChangesXml, xmlSchema);

    public string OpenHierarchy(string path, string relativeToObjectId, int createFileType)
        => GuardReturn(nameof(OpenHierarchy), path, relativeToObjectId, createFileType);

    public void Publish(string hierarchyId, string targetFilePath, int publishFormat, string clsidExporter = "")
        => GuardVoid(nameof(Publish), hierarchyId, targetFilePath, publishFormat, clsidExporter);

    public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault)
        => GuardReturn(nameof(CreateNewPage), sectionId, newPageStyle);

    public void CloseNotebook(string notebookId)
        => GuardVoid(nameof(CloseNotebook), notebookId);

    public void DeleteHierarchy(string objectId)
        => GuardVoid(nameof(DeleteHierarchy), objectId);

    public void UpdateHierarchy(string changesXml, int xmlSchema)
        => GuardVoid(nameof(UpdateHierarchy), changesXml, xmlSchema);

    public string GetBinaryPageContent(string pageId, string callbackId)
        => GuardReturn(nameof(GetBinaryPageContent), pageId, callbackId);

    public void DeletePageContent(string pageId, string objectId)
        => GuardVoid(nameof(DeletePageContent), pageId, objectId);

    public string GetHierarchyParent(string objectId)
        => GuardReturn(nameof(GetHierarchyParent), objectId);

    public string GetSpecialLocation(int specialLocation)
        => GuardReturn(nameof(GetSpecialLocation), specialLocation);

    public void NavigateTo(string hierarchyObjectId, string objectId = "", bool newWindow = false)
        => GuardVoid(nameof(NavigateTo), hierarchyObjectId, objectId, newWindow);

    public string GetHyperlinkToObject(string hierarchyId, string pageContentObjectId = "")
        => GuardReturn(nameof(GetHyperlinkToObject), hierarchyId, pageContentObjectId);

    public string FindMeta(string searchName, int xmlSchema, string startNodeId = "")
        => GuardReturn(nameof(FindMeta), searchName, xmlSchema, startNodeId);

    public void NavigateToUrl(string url, bool newWindow = false)
        => GuardVoid(nameof(NavigateToUrl), url, newWindow);

    public string GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId = "")
        => GuardReturn(nameof(GetWebHyperlinkToObject), hierarchyId, pageContentObjectId);

    public void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile)
        => GuardVoid(nameof(MergeFiles), baseFile, clientFile, serverFile, targetFile);

    public void MergeSections(string sourceId, string destId)
        => GuardVoid(nameof(MergeSections), sourceId, destId);

    public void SyncHierarchy(string hierarchyId)
        => GuardVoid(nameof(SyncHierarchy), hierarchyId);

    public void SetFilingLocation(int filingLocation, int filingLocationType, string sectionId)
        => GuardVoid(nameof(SetFilingLocation), filingLocation, filingLocationType, sectionId);
}
