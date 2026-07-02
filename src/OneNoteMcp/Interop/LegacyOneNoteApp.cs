extern alias Legacy;

using System.Runtime.Versioning;
using L = Legacy::Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp.Interop;

/// <summary>
/// Adapter over the legacy v12 (2007) OneNote COM <see cref="L.IApplication"/>.
/// The v12 API is schema-locked to the 2007 XML schema, so the <c>xmlSchema</c>
/// parameters are dropped, and six modern-only members throw
/// <see cref="NotSupportedInVersionException"/> without touching the COM object.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class LegacyOneNoteApp : IOneNoteApp
{
    private readonly L.IApplication _app;

    public LegacyOneNoteApp(L.IApplication app) => _app = app;

    /// <inheritdoc/>
    public string GetHierarchy(string startNodeId, int scope, int xmlSchema)
    {
        _app.GetHierarchy(startNodeId, (L.HierarchyScope)scope, out var xml);
        return xml;
    }

    /// <inheritdoc/>
    public string FindPages(string searchString, int xmlSchema, string startNodeId = "")
    {
        _app.FindPages(startNodeId, searchString, out var xml, false, false);
        return xml;
    }

    /// <inheritdoc/>
    public string GetPageContent(string pageId, int pageInfo, int xmlSchema)
    {
        _app.GetPageContent(pageId, out var xml, (L.PageInfo)pageInfo);
        return xml;
    }

    /// <inheritdoc/>
    public void UpdatePageContent(string pageChangesXml, int xmlSchema)
        => _app.UpdatePageContent(pageChangesXml, DateTime.MinValue);

    /// <inheritdoc/>
    public string OpenHierarchy(string path, string relativeToObjectId, int createFileType)
    {
        _app.OpenHierarchy(path, relativeToObjectId, out var id, (L.CreateFileType)createFileType);
        return id;
    }

    /// <inheritdoc/>
    public void Publish(string hierarchyId, string targetFilePath, int publishFormat, string clsidExporter = "")
        => _app.Publish(hierarchyId, targetFilePath, (L.PublishFormat)publishFormat, clsidExporter);

    /// <inheritdoc/>
    public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault)
    {
        _app.CreateNewPage(sectionId, out var id, (L.NewPageStyle)newPageStyle);
        return id;
    }

    /// <inheritdoc/>
    public void CloseNotebook(string notebookId)
        => _app.CloseNotebook(notebookId);

    /// <inheritdoc/>
    public void DeleteHierarchy(string objectId)
        => _app.DeleteHierarchy(objectId, DateTime.MinValue);

    /// <inheritdoc/>
    public void UpdateHierarchy(string changesXml, int xmlSchema)
        => _app.UpdateHierarchy(changesXml);

    /// <inheritdoc/>
    public string GetBinaryPageContent(string pageId, string callbackId)
    {
        _app.GetBinaryPageContent(pageId, callbackId, out var b64);
        return b64;
    }

    /// <inheritdoc/>
    public void DeletePageContent(string pageId, string objectId)
        => _app.DeletePageContent(pageId, objectId, DateTime.MinValue);

    /// <inheritdoc/>
    public string GetHierarchyParent(string objectId)
    {
        _app.GetHierarchyParent(objectId, out var parentId);
        return parentId;
    }

    /// <inheritdoc/>
    public string GetSpecialLocation(int specialLocation)
    {
        _app.GetSpecialLocation((L.SpecialLocation)specialLocation, out var path);
        return path;
    }

    /// <inheritdoc/>
    public void NavigateTo(string hierarchyObjectId, string objectId = "", bool newWindow = false)
        => _app.NavigateTo(hierarchyObjectId, objectId, newWindow);

    /// <inheritdoc/>
    public string GetHyperlinkToObject(string hierarchyId, string pageContentObjectId = "")
    {
        _app.GetHyperlinkToObject(hierarchyId, pageContentObjectId, out var link);
        return link;
    }

    /// <inheritdoc/>
    public string FindMeta(string searchName, int xmlSchema, string startNodeId = "")
    {
        _app.FindMeta(startNodeId, searchName, out var xml, false);
        return xml;
    }

    /// <inheritdoc/>
    public void NavigateToUrl(string url, bool newWindow = false)
        => throw new NotSupportedInVersionException(12, nameof(NavigateToUrl));

    /// <inheritdoc/>
    public string GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId = "")
        => throw new NotSupportedInVersionException(12, nameof(GetWebHyperlinkToObject));

    /// <inheritdoc/>
    public void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile)
        => throw new NotSupportedInVersionException(12, nameof(MergeFiles));

    /// <inheritdoc/>
    public void MergeSections(string sourceId, string destId)
        => throw new NotSupportedInVersionException(12, nameof(MergeSections));

    /// <inheritdoc/>
    public void SyncHierarchy(string hierarchyId)
        => throw new NotSupportedInVersionException(12, nameof(SyncHierarchy));

    /// <inheritdoc/>
    public void SetFilingLocation(int filingLocation, int filingLocationType, string sectionId)
        => throw new NotSupportedInVersionException(12, nameof(SetFilingLocation));
}
