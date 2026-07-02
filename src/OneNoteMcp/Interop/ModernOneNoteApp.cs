using System.Runtime.Versioning;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp.Interop;

/// <summary>
/// Adapter over the modern (v14+/2010+) OneNote COM <see cref="IApplication"/>.
/// Every <see cref="IOneNoteApp"/> member is supported and delegates straight through,
/// casting the schema-agnostic int parameters to the modern interop enums.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class ModernOneNoteApp : IOneNoteApp
{
    private readonly IApplication _app;

    public ModernOneNoteApp(IApplication app) => _app = app;

    /// <inheritdoc/>
    public string GetHierarchy(string startNodeId, int scope, int xmlSchema)
    {
        _app.GetHierarchy(startNodeId, (HierarchyScope)scope, out var xml, (XMLSchema)xmlSchema);
        return xml;
    }

    /// <inheritdoc/>
    public string FindPages(string searchString, int xmlSchema, string startNodeId = "")
    {
        _app.FindPages(startNodeId, searchString, out var xml, false, false, (XMLSchema)xmlSchema);
        return xml;
    }

    /// <inheritdoc/>
    public string GetPageContent(string pageId, int pageInfo, int xmlSchema)
    {
        _app.GetPageContent(pageId, out var xml, (PageInfo)pageInfo, (XMLSchema)xmlSchema);
        return xml;
    }

    /// <inheritdoc/>
    public void UpdatePageContent(string pageChangesXml, int xmlSchema)
        => _app.UpdatePageContent(pageChangesXml, DateTime.MinValue, (XMLSchema)xmlSchema, true);

    /// <inheritdoc/>
    public string OpenHierarchy(string path, string relativeToObjectId, int createFileType)
    {
        _app.OpenHierarchy(path, relativeToObjectId, out var id, (CreateFileType)createFileType);
        return id;
    }

    /// <inheritdoc/>
    public void Publish(string hierarchyId, string targetFilePath, int publishFormat, string clsidExporter = "")
        => _app.Publish(hierarchyId, targetFilePath, (PublishFormat)publishFormat, clsidExporter);

    /// <inheritdoc/>
    public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault)
    {
        _app.CreateNewPage(sectionId, out var id, (NewPageStyle)newPageStyle);
        return id;
    }

    /// <inheritdoc/>
    public void CloseNotebook(string notebookId)
        => _app.CloseNotebook(notebookId);

    /// <inheritdoc/>
    public void DeleteHierarchy(string objectId)
        => _app.DeleteHierarchy(objectId, DateTime.MinValue, true);

    /// <inheritdoc/>
    public void UpdateHierarchy(string changesXml, int xmlSchema)
        => _app.UpdateHierarchy(changesXml, (XMLSchema)xmlSchema);

    /// <inheritdoc/>
    public string GetBinaryPageContent(string pageId, string callbackId)
    {
        _app.GetBinaryPageContent(pageId, callbackId, out var b64);
        return b64;
    }

    /// <inheritdoc/>
    public void DeletePageContent(string pageId, string objectId)
        => _app.DeletePageContent(pageId, objectId, DateTime.MinValue, true);

    /// <inheritdoc/>
    public string GetHierarchyParent(string objectId)
    {
        _app.GetHierarchyParent(objectId, out var parentId);
        return parentId;
    }

    /// <inheritdoc/>
    public string GetSpecialLocation(int specialLocation)
    {
        _app.GetSpecialLocation((SpecialLocation)specialLocation, out var path);
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
        _app.FindMeta(startNodeId, searchName, out var xml, false, (XMLSchema)xmlSchema);
        return xml;
    }

    /// <inheritdoc/>
    public void NavigateToUrl(string url, bool newWindow = false)
        => _app.NavigateToUrl(url, newWindow);

    /// <inheritdoc/>
    public string GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId = "")
    {
        _app.GetWebHyperlinkToObject(hierarchyId, pageContentObjectId, out var link);
        return link;
    }

    /// <inheritdoc/>
    public void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile)
        => _app.MergeFiles(baseFile, clientFile, serverFile, targetFile);

    /// <inheritdoc/>
    public void MergeSections(string sourceId, string destId)
        => _app.MergeSections(sourceId, destId);

    /// <inheritdoc/>
    public void SyncHierarchy(string hierarchyId)
        => _app.SyncHierarchy(hierarchyId);

    /// <inheritdoc/>
    public void SetFilingLocation(int filingLocation, int filingLocationType, string sectionId)
        => _app.SetFilingLocation((FilingLocation)filingLocation, (FilingLocationType)filingLocationType, sectionId);
}
