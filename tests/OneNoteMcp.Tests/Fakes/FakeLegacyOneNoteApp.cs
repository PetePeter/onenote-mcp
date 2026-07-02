extern alias Legacy;
using L = Legacy::Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp.Tests.Fakes;

/// <summary>
/// COM-free recording fake of the v12 (OneNote 2007) <c>IApplication</c> surface
/// (18 methods, no XMLSchema parameter, schema-locked to the 2007 schema). Passed
/// to <c>LegacyOneNoteApp</c> to prove it maps each supported <c>IOneNoteApp</c>
/// call onto v12's shorter signature — and, critically, that the six
/// IApplication2–4 methods absent on v12 throw <c>NotSupportedInVersionException</c>
/// WITHOUT ever reaching this RCW (asserted via <see cref="LastMethod"/> staying null).
/// </summary>
public sealed class FakeLegacyOneNoteApp : L.IApplication
{
    public string? LastMethod { get; private set; }
    public List<object?> LastArgs { get; private set; } = new();
    public string ReturnValue { get; set; } = "";

    private void Record(string method, params object?[] inputs)
    {
        LastMethod = method;
        LastArgs = inputs.ToList();
    }

    public void GetHierarchy(string bstrStartNodeID, L.HierarchyScope hsScope, out string pbstrHierarchyXmlOut)
    {
        Record(nameof(GetHierarchy), bstrStartNodeID, (int)hsScope);
        pbstrHierarchyXmlOut = ReturnValue;
    }

    public void FindPages(string bstrStartNodeID, string bstrSearchString,
        out string pbstrHierarchyXmlOut, bool fIncludeUnindexedPages, bool fDisplay)
    {
        Record(nameof(FindPages), bstrStartNodeID, bstrSearchString, fIncludeUnindexedPages, fDisplay);
        pbstrHierarchyXmlOut = ReturnValue;
    }

    public void GetPageContent(string bstrPageID, out string pbstrPageXmlOut, L.PageInfo pageInfoToExport)
    {
        Record(nameof(GetPageContent), bstrPageID, (int)pageInfoToExport);
        pbstrPageXmlOut = ReturnValue;
    }

    public void UpdatePageContent(string bstrPageChangesXmlIn, DateTime dateExpectedLastModified)
        => Record(nameof(UpdatePageContent), bstrPageChangesXmlIn, dateExpectedLastModified);

    public void OpenHierarchy(string bstrPath, string bstrRelativeToObjectID,
        out string pbstrObjectID, L.CreateFileType cftIfNotExist)
    {
        Record(nameof(OpenHierarchy), bstrPath, bstrRelativeToObjectID, (int)cftIfNotExist);
        pbstrObjectID = ReturnValue;
    }

    public void Publish(string bstrHierarchyID, string bstrTargetFilePath,
        L.PublishFormat pfPublishFormat, string bstrCLSIDofExporter)
        => Record(nameof(Publish), bstrHierarchyID, bstrTargetFilePath, (int)pfPublishFormat, bstrCLSIDofExporter);

    public void CreateNewPage(string bstrSectionID, out string pbstrPageID, L.NewPageStyle npsNewPageStyle)
    {
        Record(nameof(CreateNewPage), bstrSectionID, (int)npsNewPageStyle);
        pbstrPageID = ReturnValue;
    }

    public void CloseNotebook(string bstrNotebookID)
        => Record(nameof(CloseNotebook), bstrNotebookID);

    public void DeleteHierarchy(string bstrObjectID, DateTime dateExpectedLastModified)
        => Record(nameof(DeleteHierarchy), bstrObjectID, dateExpectedLastModified);

    public void UpdateHierarchy(string bstrChangesXmlIn)
        => Record(nameof(UpdateHierarchy), bstrChangesXmlIn);

    public void GetBinaryPageContent(string bstrPageID, string bstrCallbackID, out string pbstrBinaryObjectB64Out)
    {
        Record(nameof(GetBinaryPageContent), bstrPageID, bstrCallbackID);
        pbstrBinaryObjectB64Out = ReturnValue;
    }

    public void DeletePageContent(string bstrPageID, string bstrObjectID, DateTime dateExpectedLastModified)
        => Record(nameof(DeletePageContent), bstrPageID, bstrObjectID, dateExpectedLastModified);

    public void GetHierarchyParent(string bstrObjectID, out string pbstrParentID)
    {
        Record(nameof(GetHierarchyParent), bstrObjectID);
        pbstrParentID = ReturnValue;
    }

    public void GetSpecialLocation(L.SpecialLocation slToGet, out string pbstrSpecialLocationPath)
    {
        Record(nameof(GetSpecialLocation), (int)slToGet);
        pbstrSpecialLocationPath = ReturnValue;
    }

    public void NavigateTo(string bstrHierarchyObjectID, string bstrObjectID, bool fNewWindow)
        => Record(nameof(NavigateTo), bstrHierarchyObjectID, bstrObjectID, fNewWindow);

    public void GetHyperlinkToObject(string bstrHierarchyID, string bstrPageContentObjectID, out string pbstrHyperlinkOut)
    {
        Record(nameof(GetHyperlinkToObject), bstrHierarchyID, bstrPageContentObjectID);
        pbstrHyperlinkOut = ReturnValue;
    }

    public void FindMeta(string bstrStartNodeID, string bstrSearchStringName,
        out string pbstrHierarchyXmlOut, bool fIncludeUnindexedPages)
    {
        Record(nameof(FindMeta), bstrStartNodeID, bstrSearchStringName, fIncludeUnindexedPages);
        pbstrHierarchyXmlOut = ReturnValue;
    }

    // v12 exposes OpenPackage; the adapters don't map it — fail loudly if called.
    public void OpenPackage(string bstrPathPackage, string bstrPathDest, out string pbstrPathOut)
        => throw new NotImplementedException();
}
