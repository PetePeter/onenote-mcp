extern alias Legacy;
using System.Collections.Concurrent;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteMcp.Interop;

// ─── Scope / schema / format constants ────────────────────────────────────────

/// <summary>OneNote HierarchyScope enumeration values.</summary>
public static class OneNoteScope
{
    public const int HsSelf      = 0;
    public const int HsChildren  = 1;
    public const int HsNotebooks = 2;
    public const int HsSections  = 3;
    public const int HsPages     = 4;
}

/// <summary>OneNote XMLSchema enumeration values.</summary>
public static class OneNoteXmlSchema
{
    public const int Xs2007 = 0;
    public const int Xs2010 = 1;
    public const int Xs2013 = 2;
}

/// <summary>OneNote CreateFileType enumeration values.</summary>
public static class OneNoteCreateFileType
{
    public const int CftNone     = 0;
    public const int CftNotebook = 1;
    public const int CftFolder   = 2;
    public const int CftSection  = 3;
}

/// <summary>OneNote NewPageStyle enumeration values.</summary>
public static class OneNoteNewPageStyle
{
    public const int NpsDefault            = 0;
    public const int NpsBlankPageWithTitle = 1;
    public const int NpsBlankPageNoTitle   = 2;
}

/// <summary>OneNote PublishFormat enumeration values.</summary>
public static class OneNotePublishFormat
{
    public const int PfOneNote        = 0;
    public const int PfOneNotePackage = 1;
    public const int PfMhtml          = 2;
    public const int PfPdf            = 3;
    public const int PfXps            = 4;
    public const int PfWord           = 5;
    public const int PfEmf            = 6;
    public const int PfHtml           = 7;
    public const int PfOneNote2007    = 8;
}

/// <summary>OneNote SpecialLocation enumeration values (for GetSpecialLocation).</summary>
public static class OneNoteSpecialLocation
{
    public const int SlBackUpFolder          = 0;
    public const int SlUnfiledNotesSection   = 1;
    public const int SlDefaultNotebookFolder = 2;
}

/// <summary>OneNote FilingLocation enumeration values (for SetFilingLocation).</summary>
public static class OneNoteFilingLocation
{
    public const int FlEMail      = 0;
    public const int FlContacts   = 1;
    public const int FlTasks      = 2;
    public const int FlMeetings   = 3;
    public const int FlWebContent = 4;
    public const int FlPrintOuts  = 5;
}

/// <summary>
/// OneNote FilingLocationType enumeration values (for SetFilingLocation).
/// Note the interop skips 3 — <c>fltNamedPage</c> is 4.
/// </summary>
public static class OneNoteFilingLocationType
{
    public const int FltNamedSectionNewPage  = 0;
    public const int FltCurrentSectionNewPage = 1;
    public const int FltCurrentPage          = 2;
    public const int FltNamedPage            = 4;
}

// ─── Session ──────────────────────────────────────────────────────────────────

/// <summary>
/// Lazy, recreatable bridge to the OneNote COM Application object, keyed per CLSID.
/// All COM calls are serialised through a single semaphore because the OneNote
/// Application object lives in a single-threaded COM apartment.
/// Use <see cref="For(string)"/> to obtain a per-CLSID session; do NOT use a
/// process-wide singleton — different OneNote versions require different CLSIDs.
/// </summary>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public sealed class OneNoteSession : IDisposable
{
    // HRESULT for "RPC server unavailable" — COM object must be recreated
    private const int HrServerUnavailable = unchecked((int)0x800706BA);

    // ── Per-CLSID session cache ───────────────────────────────────────────────

    private static readonly ConcurrentDictionary<string, OneNoteSession> _sessions =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Test seam: when set, replaces the real COM factory for every new session.
    /// The override receives the CLSID string and returns an <see cref="IOneNoteApp"/>.
    /// </summary>
    internal static Func<string, IOneNoteApp>? AppFactoryOverride { get; set; }

    /// <summary>
    /// Returns (or creates) a session for the given canonical CLSID string.
    /// Sessions are cached process-wide: the same CLSID always returns the same
    /// instance so COM objects are not multiplied.
    /// </summary>
    public static OneNoteSession For(string clsid) =>
        _sessions.GetOrAdd(clsid, c => new OneNoteSession(() => CreateApp(c)));

    /// <summary>
    /// Disposes all cached sessions, clears the cache, and nulls the test override.
    /// Must be called in test setup and teardown to prevent state leaks between tests.
    /// </summary>
    internal static void ResetForTests()
    {
        foreach (var s in _sessions.Values) s.Dispose();
        _sessions.Clear();
        AppFactoryOverride = null;
    }

    private static IOneNoteApp CreateApp(string clsid) =>
        AppFactoryOverride?.Invoke(clsid) ?? CreateRealApp(clsid);

    /// <summary>
    /// Instantiates a COM server by CLSID and wraps it in the appropriate adapter.
    /// Major 12 (OneNote 2007) gets a LegacyOneNoteApp; all other versions use the
    /// modern adapter which supports the full 23-method surface.
    /// </summary>
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static IOneNoteApp CreateRealApp(string clsid)
    {
        var guid = new Guid(clsid);
        var canonical = guid.ToString("B").ToUpperInvariant();
        var major = OneNoteVersionCatalog.All
            .FirstOrDefault(k => string.Equals(k.Clsid, canonical, StringComparison.OrdinalIgnoreCase))
            ?.Major ?? 16;

        var type = Type.GetTypeFromCLSID(guid)
            ?? throw new InvalidOperationException($"No COM class registered for CLSID {clsid}.");
        var rcw = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException($"Failed to instantiate OneNote COM server {clsid}.");

        return major == 12
            ? new LegacyOneNoteApp((Legacy::Microsoft.Office.Interop.OneNote.IApplication)rcw)
            : new ModernOneNoteApp((IApplication)rcw);
    }

    // ── Instance state ────────────────────────────────────────────────────────

    /// <summary>
    /// The most recent COM failure (mapped to human-readable text) seen by any
    /// session in this process. Null until the first COM failure.
    /// </summary>
    public static string? LastComError { get; private set; }

    private readonly Func<IOneNoteApp> _appFactory;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private IOneNoteApp? _app;

    /// <summary>
    /// Creates a session with an injectable COM app factory.
    /// Used by tests to verify recreate-on-death behaviour without real COM.
    /// </summary>
    internal OneNoteSession(Func<IOneNoteApp> appFactory)
    {
        _appFactory = appFactory;
    }

    /// <summary>
    /// True when the OneNote ProgID is registered on this machine, meaning
    /// OneNote is installed and COM interop is possible.
    /// </summary>
    public static bool IsComAvailable =>
        Type.GetTypeFromProgID("OneNote.Application") is not null;

    /// <summary>
    /// Human-readable version string from registry detection, e.g. "2016/2019/365".
    /// Null when OneNote is not detected.
    /// </summary>
    public string? DetectedVersionDisplay => OneNoteVersion.Detect()?.DisplayName;

    // ── Internal machinery ────────────────────────────────────────────────────

    /// <summary>
    /// Ensures the COM application object exists, (re)creating it via the
    /// factory when it is null (first use or after a dead-proxy invalidation).
    /// Must be called inside the semaphore.
    /// </summary>
    private void EnsureApp()
    {
        if (_app is null)
            _app = _appFactory();
    }

    /// <summary>
    /// Marks the current COM app object as dead so <see cref="EnsureApp"/>
    /// will recreate it on the next call.
    /// </summary>
    private void InvalidateApp() => _app = null;

    /// <summary>Exposes invalidation for test scenarios that simulate a dead COM proxy.</summary>
    internal void InvalidateForTests() => InvalidateApp();

    /// <summary>
    /// Runs a COM operation under the serialising semaphore with retry and
    /// recreate-on-death. Delegates through the IOneNoteApp adapter so the same
    /// path works for both the modern and legacy (v12) COM servers.
    /// </summary>
    private T Dispatch<T>(Func<IOneNoteApp, T> op, int maxAttempts = 5)
    {
        _lock.Wait();
        try
        {
            return ComRetry.Execute(() =>
            {
                EnsureApp();
                try
                {
                    return op(_app!);
                }
                catch (COMException ex)
                {
                    // Record every COM failure for diagnostics before existing handling.
                    LastComError = ComErrorMapper.Describe(ex);
                    if (ex.HResult == HrServerUnavailable)
                        InvalidateApp(); // Dead proxy — recreate on next call.
                    throw;
                }
            }, maxAttempts);
        }
        catch (COMException ex) when (ex.HResult == HrServerUnavailable)
        {
            // Guard: ensure invalidation even if ComRetry rethrew server-unavailable.
            InvalidateApp();
            throw;
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>Void overload of <see cref="Dispatch{T}"/>.</summary>
    private void Dispatch(Action<IOneNoteApp> op) =>
        Dispatch<int>(a => { op(a); return 0; });

    /// <summary>
    /// Dispatches a NON-idempotent operation with retries disabled (maxAttempts=1).
    /// OneNote's STA COM frequently returns transient RPC_E_CALL_REJECTED /
    /// RPC_E_SERVERCALL_RETRYLATER <em>after</em> a create/new has already taken
    /// effect; retrying such an op would duplicate the created object (e.g. a page).
    /// The serialising semaphore and recreate-on-death behaviour are preserved; only
    /// the transient-error retry loop is suppressed.
    /// </summary>
    private T DispatchOnce<T>(Func<IOneNoteApp, T> op) => Dispatch(op, maxAttempts: 1);

    // ── Public COM wrappers ───────────────────────────────────────────────────
    // Each method delegates straight to IOneNoteApp via Dispatch.

    /// <summary>
    /// Returns the OneNote hierarchy XML starting from <paramref name="startNodeId"/>.
    /// Pass empty string for the root (all notebooks).
    /// </summary>
    public string GetHierarchy(string startNodeId, int scope, int xmlSchema) =>
        Dispatch(a => a.GetHierarchy(startNodeId, scope, xmlSchema));

    /// <summary>Runs a OneNote search and returns hierarchy XML of matching pages.</summary>
    public string FindPages(string searchString, int xmlSchema, string startNodeId = "") =>
        Dispatch(a => a.FindPages(searchString, xmlSchema, startNodeId));

    /// <summary>Returns the XML content of a OneNote page.</summary>
    public string GetPageContent(string pageId, int pageInfo, int xmlSchema) =>
        Dispatch(a => a.GetPageContent(pageId, pageInfo, xmlSchema));

    /// <summary>Saves modified page XML back to OneNote.</summary>
    public void UpdatePageContent(string pageChangesXml, int xmlSchema) =>
        Dispatch(a => a.UpdatePageContent(pageChangesXml, xmlSchema));

    /// <summary>
    /// Opens or creates a hierarchy node. Returns the object ID of the node.
    /// </summary>
    public string OpenHierarchy(string path, string relativeToObjectId, int createFileType) =>
        Dispatch(a => a.OpenHierarchy(path, relativeToObjectId, createFileType));

    /// <summary>Publishes a hierarchy node to a file in the given format.</summary>
    public void Publish(string hierarchyId, string targetFilePath, int publishFormat,
        string clsidExporter = "") =>
        Dispatch(a => a.Publish(hierarchyId, targetFilePath, publishFormat, clsidExporter));

    /// <summary>Creates a new page in a section. Returns the new page's object ID.</summary>
    public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault) =>
        DispatchOnce(a => a.CreateNewPage(sectionId, newPageStyle));

    /// <summary>Closes an open notebook. Does not delete files on disk.</summary>
    public void CloseNotebook(string notebookId) =>
        Dispatch(a => a.CloseNotebook(notebookId));

    /// <summary>Deletes a hierarchy node (page/section/notebook) by its object ID. No conflict check.</summary>
    public void DeleteHierarchy(string objectId) =>
        Dispatch(a => a.DeleteHierarchy(objectId));

    /// <summary>Applies a hierarchy-change XML fragment (e.g. a renamed node) to OneNote.</summary>
    public void UpdateHierarchy(string changesXml, int xmlSchema) =>
        Dispatch(a => a.UpdateHierarchy(changesXml, xmlSchema));

    /// <summary>
    /// Fetches a page's binary object (identified by a callback ID from the page
    /// XML) as a base64 string.
    /// </summary>
    public string GetBinaryPageContent(string pageId, string callbackId) =>
        Dispatch(a => a.GetBinaryPageContent(pageId, callbackId));

    /// <summary>Deletes a single content object (by object ID) from a page.</summary>
    public void DeletePageContent(string pageId, string objectId) =>
        Dispatch(a => a.DeletePageContent(pageId, objectId));

    /// <summary>Returns the object ID of the parent of the given hierarchy node.</summary>
    public string GetHierarchyParent(string objectId) =>
        Dispatch(a => a.GetHierarchyParent(objectId));

    /// <summary>
    /// Returns the filesystem path of a OneNote special location
    /// (backup folder, unfiled-notes section, or default notebook folder).
    /// </summary>
    public string GetSpecialLocation(int specialLocation) =>
        Dispatch(a => a.GetSpecialLocation(specialLocation));

    /// <summary>Navigates the OneNote UI to a hierarchy node and optional object.</summary>
    public void NavigateTo(string hierarchyObjectId, string objectId = "", bool newWindow = false) =>
        Dispatch(a => a.NavigateTo(hierarchyObjectId, objectId, newWindow));

    /// <summary>Navigates the OneNote UI to a onenote: URL.</summary>
    public void NavigateToUrl(string url, bool newWindow = false) =>
        Dispatch(a => a.NavigateToUrl(url, newWindow));

    /// <summary>Returns a onenote: hyperlink to a hierarchy node (and optional page object).</summary>
    public string GetHyperlinkToObject(string hierarchyId, string pageContentObjectId = "") =>
        Dispatch(a => a.GetHyperlinkToObject(hierarchyId, pageContentObjectId));

    /// <summary>Returns a web (https) hyperlink to a hierarchy node (and optional page object).</summary>
    public string GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId = "") =>
        Dispatch(a => a.GetWebHyperlinkToObject(hierarchyId, pageContentObjectId));

    /// <summary>
    /// Searches page metadata by name and returns hierarchy XML of matching pages.
    /// Pass empty <paramref name="startNodeId"/> to search from the root.
    /// </summary>
    public string FindMeta(string searchName, int xmlSchema, string startNodeId = "") =>
        Dispatch(a => a.FindMeta(searchName, xmlSchema, startNodeId));

    /// <summary>Three-way merges OneNote files (base/client/server) into a target file.</summary>
    public void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile) =>
        Dispatch(a => a.MergeFiles(baseFile, clientFile, serverFile, targetFile));

    /// <summary>Merges the pages of a source section into a destination section.</summary>
    public void MergeSections(string sourceId, string destId) =>
        Dispatch(a => a.MergeSections(sourceId, destId));

    /// <summary>Forces a sync of the given hierarchy node (notebook/section).</summary>
    public void SyncHierarchy(string hierarchyId) =>
        Dispatch(a => a.SyncHierarchy(hierarchyId));

    /// <summary>
    /// Sets the section OneNote files a given kind of Outlook item into
    /// (e-mail, contacts, tasks, …).
    /// </summary>
    public void SetFilingLocation(int filingLocation, int filingLocationType, string sectionId) =>
        Dispatch(a => a.SetFilingLocation(filingLocation, filingLocationType, sectionId));

    /// <inheritdoc/>
    public void Dispose() => _lock.Dispose();
}
