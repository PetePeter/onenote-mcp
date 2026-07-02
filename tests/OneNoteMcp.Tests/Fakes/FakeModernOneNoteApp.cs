using Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp.Tests.Fakes;

/// <summary>
/// COM-free recording fake of the MODERN OneNote <c>IApplication</c> surface
/// (the interface embedded into OneNoteMcp.dll). Passed to
/// <c>ModernOneNoteApp</c> to prove it delegates each <c>IOneNoteApp</c> call to
/// the matching typed RCW method with the right positional arguments, and writes
/// back the canned out-value. Enum arguments are recorded as their underlying
/// <see cref="int"/> so schema/format pass-through can be asserted.
///
/// Members the adapters never touch (properties, QuickFiling, OpenPackage) throw
/// so an accidental call fails loudly rather than passing silently.
/// </summary>
public sealed class FakeModernOneNoteApp : IApplication
{
    public string? LastMethod { get; private set; }
    public List<object?> LastArgs { get; private set; } = new();
    public string ReturnValue { get; set; } = "";

    private void Record(string method, params object?[] inputs)
    {
        LastMethod = method;
        LastArgs = inputs.ToList();
    }

    public void GetHierarchy(string bstrStartNodeID, HierarchyScope hsScope,
        out string pbstrHierarchyXmlOut, XMLSchema xsSchema)
    {
        Record(nameof(GetHierarchy), bstrStartNodeID, (int)hsScope, (int)xsSchema);
        pbstrHierarchyXmlOut = ReturnValue;
    }

    public void FindPages(string bstrStartNodeID, string bstrSearchString,
        out string pbstrHierarchyXmlOut, bool fIncludeUnindexedPages, bool fDisplay, XMLSchema xsSchema)
    {
        Record(nameof(FindPages), bstrStartNodeID, bstrSearchString, fIncludeUnindexedPages, fDisplay, (int)xsSchema);
        pbstrHierarchyXmlOut = ReturnValue;
    }

    public void GetPageContent(string bstrPageID, out string pbstrPageXmlOut,
        PageInfo pageInfoToExport, XMLSchema xsSchema)
    {
        Record(nameof(GetPageContent), bstrPageID, (int)pageInfoToExport, (int)xsSchema);
        pbstrPageXmlOut = ReturnValue;
    }

    public void UpdatePageContent(string bstrPageChangesXmlIn, DateTime dateExpectedLastModified,
        XMLSchema xsSchema, bool force)
        => Record(nameof(UpdatePageContent), bstrPageChangesXmlIn, dateExpectedLastModified, (int)xsSchema, force);

    public void OpenHierarchy(string bstrPath, string bstrRelativeToObjectID,
        out string pbstrObjectID, CreateFileType cftIfNotExist)
    {
        Record(nameof(OpenHierarchy), bstrPath, bstrRelativeToObjectID, (int)cftIfNotExist);
        pbstrObjectID = ReturnValue;
    }

    public void Publish(string bstrHierarchyID, string bstrTargetFilePath,
        PublishFormat pfPublishFormat, string bstrCLSIDofExporter)
        => Record(nameof(Publish), bstrHierarchyID, bstrTargetFilePath, (int)pfPublishFormat, bstrCLSIDofExporter);

    public void CreateNewPage(string bstrSectionID, out string pbstrPageID, NewPageStyle npsNewPageStyle)
    {
        Record(nameof(CreateNewPage), bstrSectionID, (int)npsNewPageStyle);
        pbstrPageID = ReturnValue;
    }

    public void CloseNotebook(string bstrNotebookID, bool force)
        => Record(nameof(CloseNotebook), bstrNotebookID, force);

    public void DeleteHierarchy(string bstrObjectID, DateTime dateExpectedLastModified, bool deletePermanently)
        => Record(nameof(DeleteHierarchy), bstrObjectID, dateExpectedLastModified, deletePermanently);

    public void UpdateHierarchy(string bstrChangesXmlIn, XMLSchema xsSchema)
        => Record(nameof(UpdateHierarchy), bstrChangesXmlIn, (int)xsSchema);

    public void GetBinaryPageContent(string bstrPageID, string bstrCallbackID, out string pbstrBinaryObjectB64Out)
    {
        Record(nameof(GetBinaryPageContent), bstrPageID, bstrCallbackID);
        pbstrBinaryObjectB64Out = ReturnValue;
    }

    public void DeletePageContent(string bstrPageID, string bstrObjectID,
        DateTime dateExpectedLastModified, bool force)
        => Record(nameof(DeletePageContent), bstrPageID, bstrObjectID, dateExpectedLastModified, force);

    public void GetHierarchyParent(string bstrObjectID, out string pbstrParentID)
    {
        Record(nameof(GetHierarchyParent), bstrObjectID);
        pbstrParentID = ReturnValue;
    }

    public void GetSpecialLocation(SpecialLocation slToGet, out string pbstrSpecialLocationPath)
    {
        Record(nameof(GetSpecialLocation), (int)slToGet);
        pbstrSpecialLocationPath = ReturnValue;
    }

    public void NavigateTo(string bstrHierarchyObjectID, string bstrObjectID, bool fNewWindow)
        => Record(nameof(NavigateTo), bstrHierarchyObjectID, bstrObjectID, fNewWindow);

    public void NavigateToUrl(string bstrUrl, bool fNewWindow)
        => Record(nameof(NavigateToUrl), bstrUrl, fNewWindow);

    public void GetHyperlinkToObject(string bstrHierarchyID, string bstrPageContentObjectID, out string pbstrHyperlinkOut)
    {
        Record(nameof(GetHyperlinkToObject), bstrHierarchyID, bstrPageContentObjectID);
        pbstrHyperlinkOut = ReturnValue;
    }

    public void GetWebHyperlinkToObject(string bstrHierarchyID, string bstrPageContentObjectID, out string pbstrHyperlinkOut)
    {
        Record(nameof(GetWebHyperlinkToObject), bstrHierarchyID, bstrPageContentObjectID);
        pbstrHyperlinkOut = ReturnValue;
    }

    public void FindMeta(string bstrStartNodeID, string bstrSearchStringName,
        out string pbstrHierarchyXmlOut, bool fIncludeUnindexedPages, XMLSchema xsSchema)
    {
        Record(nameof(FindMeta), bstrStartNodeID, bstrSearchStringName, fIncludeUnindexedPages, (int)xsSchema);
        pbstrHierarchyXmlOut = ReturnValue;
    }

    public void MergeFiles(string bstrBaseFile, string bstrClientFile, string bstrServerFile, string bstrTargetFile)
        => Record(nameof(MergeFiles), bstrBaseFile, bstrClientFile, bstrServerFile, bstrTargetFile);

    public void MergeSections(string bstrSectionSourceId, string bstrSectionDestinationId)
        => Record(nameof(MergeSections), bstrSectionSourceId, bstrSectionDestinationId);

    public void SyncHierarchy(string bstrHierarchyID)
        => Record(nameof(SyncHierarchy), bstrHierarchyID);

    public void SetFilingLocation(FilingLocation flToSet, FilingLocationType fltToSet, string bstrFilingSectionID)
        => Record(nameof(SetFilingLocation), (int)flToSet, (int)fltToSet, bstrFilingSectionID);

    // ── Members the adapters do not use — fail loudly if ever invoked. ─────────
    public void OpenPackage(string bstrPathPackage, string bstrPathDest, out string pbstrPathOut)
        => throw new NotImplementedException();

    public IQuickFilingDialog QuickFiling() => throw new NotImplementedException();

    public Windows Windows => throw new NotImplementedException();
    public bool Dummy1 => throw new NotImplementedException();
    public object COMAddIns => throw new NotImplementedException();
    public object LanguageSettings => throw new NotImplementedException();
}
