namespace OneNoteMcp.Interop;

/// <summary>
/// Version-agnostic abstraction over the OneNote <c>IApplication</c> COM surface.
/// One typed implementation per supported OneNote generation:
/// <see cref="ModernOneNoteApp"/> (2010+, full 23-method surface) and
/// <see cref="LegacyOneNoteApp"/> (2007/v12, 17 methods present, the 6
/// IApplication2–4 additions throw <see cref="NotSupportedInVersionException"/>).
///
/// Signatures mirror the public wrappers on <see cref="OneNoteSession"/>
/// (string in/out, int constants from the OneNote* constant classes). The
/// <paramref name="xmlSchema"/>-style ints are honoured by the modern adapter and
/// dropped by the legacy adapter, which is schema-locked to the 2007 schema.
/// </summary>
public interface IOneNoteApp
{
    // ── Shared surface (present on every supported version) ────────────────────

    /// <summary>Returns hierarchy XML from <paramref name="startNodeId"/> (empty = root).</summary>
    string GetHierarchy(string startNodeId, int scope, int xmlSchema);

    /// <summary>Runs a search and returns hierarchy XML of matching pages.</summary>
    string FindPages(string searchString, int xmlSchema, string startNodeId = "");

    /// <summary>Returns the XML content of a page.</summary>
    string GetPageContent(string pageId, int pageInfo, int xmlSchema);

    /// <summary>Saves modified page XML back to OneNote.</summary>
    void UpdatePageContent(string pageChangesXml, int xmlSchema);

    /// <summary>Opens or creates a hierarchy node; returns its object ID.</summary>
    string OpenHierarchy(string path, string relativeToObjectId, int createFileType);

    /// <summary>Publishes a hierarchy node to a file in the given format.</summary>
    void Publish(string hierarchyId, string targetFilePath, int publishFormat, string clsidExporter = "");

    /// <summary>Creates a new page in a section; returns the new page's object ID.</summary>
    string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault);

    /// <summary>Closes an open notebook (does not delete files on disk).</summary>
    void CloseNotebook(string notebookId);

    /// <summary>Deletes a hierarchy node by object ID.</summary>
    void DeleteHierarchy(string objectId);

    /// <summary>Applies a hierarchy-change XML fragment to OneNote.</summary>
    void UpdateHierarchy(string changesXml, int xmlSchema);

    /// <summary>Fetches a page's binary object (by callback ID) as a base64 string.</summary>
    string GetBinaryPageContent(string pageId, string callbackId);

    /// <summary>Deletes a single content object (by object ID) from a page.</summary>
    void DeletePageContent(string pageId, string objectId);

    /// <summary>Returns the object ID of the parent of the given hierarchy node.</summary>
    string GetHierarchyParent(string objectId);

    /// <summary>Returns the filesystem path of a OneNote special location.</summary>
    string GetSpecialLocation(int specialLocation);

    /// <summary>Navigates the OneNote UI to a hierarchy node and optional object.</summary>
    void NavigateTo(string hierarchyObjectId, string objectId = "", bool newWindow = false);

    /// <summary>Returns a onenote: hyperlink to a hierarchy node (and optional page object).</summary>
    string GetHyperlinkToObject(string hierarchyId, string pageContentObjectId = "");

    /// <summary>Searches page metadata by name and returns hierarchy XML of matching pages.</summary>
    string FindMeta(string searchName, int xmlSchema, string startNodeId = "");

    // ── Modern-only surface (IApplication2–4; throws on v12) ───────────────────

    /// <summary>Navigates the OneNote UI to a onenote: URL. Not supported on 2007.</summary>
    void NavigateToUrl(string url, bool newWindow = false);

    /// <summary>Returns a web (https) hyperlink to a node. Not supported on 2007.</summary>
    string GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId = "");

    /// <summary>Three-way merges OneNote files. Not supported on 2007.</summary>
    void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile);

    /// <summary>Merges the pages of a source section into a destination section. Not supported on 2007.</summary>
    void MergeSections(string sourceId, string destId);

    /// <summary>Forces a sync of the given hierarchy node. Not supported on 2007.</summary>
    void SyncHierarchy(string hierarchyId);

    /// <summary>Sets the section OneNote files a kind of Outlook item into. Not supported on 2007.</summary>
    void SetFilingLocation(int filingLocation, int filingLocationType, string sectionId);
}
