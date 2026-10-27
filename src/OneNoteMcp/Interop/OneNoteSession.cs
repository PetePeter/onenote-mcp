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
/// Activation, every COM call and the final release run serially on one dedicated
/// STA thread (<see cref="StaComExecutor"/>), so the RCW has a single owner and is
/// released deterministically on <see cref="Dispose"/>.
/// Use <see cref="For(string)"/> to obtain a per-CLSID session; do NOT use a
/// process-wide singleton — different OneNote versions require different CLSIDs.
/// </summary>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public sealed class OneNoteSession : IDisposable
{
    // HRESULT for "RPC server unavailable" — COM object must be recreated
    private const int HrServerUnavailable = unchecked((int)0x800706BA);
    // HRESULT CO_E_SERVER_EXEC_FAILURE — typically OneNote is still exiting after the
    // last client released it; activation succeeds once shutdown completes.
    private const int HrServerExecFailure = unchecked((int)0x80080005);
    private const int MaxActivationAttempts = 4; // waits 1s, 2s, 4s between attempts

    /// <summary>Test seam: the wait between activation attempts. Defaults to Thread.Sleep.</summary>
    internal static Action<TimeSpan> ActivationBackoff { get; set; } = Thread.Sleep;

    // ── Per-CLSID session cache ───────────────────────────────────────────────

    private static readonly ConcurrentDictionary<string, OneNoteSession> _sessions =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Test seam: when set, replaces the real COM factory for every new session.
    /// The override receives the CLSID string and returns an <see cref="IOneNoteApp"/>.
    /// </summary>
    internal static Func<string, IOneNoteApp>? AppFactoryOverride { get; set; }

    /// <summary>Real-COM sessions evicted by <see cref="ResetForTests"/>, kept alive until process exit.</summary>
    private static readonly ConcurrentBag<OneNoteSession> _retainedForTests = new();

    /// <summary>
    /// Returns (or creates) a session for the given canonical CLSID string.
    /// Sessions are cached process-wide: the same CLSID always returns the same
    /// instance so COM objects are not multiplied.
    /// </summary>
    public static OneNoteSession For(string clsid) =>
        _sessions.GetOrAdd(clsid, c => new OneNoteSession(() => CreateApp(c)));

    /// <summary>
    /// Evicts all cached sessions (disposing fake-backed ones, parking real-COM ones)
    /// and resets the test seams. Call in test setup and teardown to prevent leaks.
    /// </summary>
    internal static void ResetForTests()
    {
        foreach (var clsid in _sessions.Keys)
        {
            if (!_sessions.TryRemove(clsid, out var session)) continue;
            // Releasing a real OneNote mid-run lets it exit and relaunch between
            // tests; park real sessions instead so they live until process exit.
            if (session._app is ModernOneNoteApp or LegacyOneNoteApp)
                _retainedForTests.Add(session);
            else
                session.Dispose();
        }
        AppFactoryOverride = null;
        ActivationBackoff = Thread.Sleep;
    }

    /// <summary>
    /// Disposes and evicts every cached session, releasing their COM objects.
    /// Called on host shutdown; a later <see cref="For"/> starts a fresh session.
    /// </summary>
    public static void DisposeAll()
    {
        foreach (var clsid in _sessions.Keys)
            if (_sessions.TryRemove(clsid, out var session))
                session.Dispose();
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
    private readonly StaComExecutor _executor = new();
    private volatile IOneNoteApp? _app;
    private volatile bool _disposed;
    private int _disposeStarted;

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
    /// Must be called on the executor thread.
    /// </summary>
    private void EnsureApp()
    {
        for (int attempt = 1; _app is null; attempt++)
        {
            try
            {
                _app = _appFactory();
            }
            catch (COMException ex) when (ex.HResult == HrServerExecFailure && attempt < MaxActivationAttempts)
            {
                ActivationBackoff(TimeSpan.FromSeconds(1 << (attempt - 1)));
            }
        }
    }

    /// <summary>
    /// Releases the current COM app object (best effort: it may already be dead)
    /// so <see cref="EnsureApp"/> recreates it on the next call.
    /// Must be called on the executor thread.
    /// </summary>
    private void InvalidateApp()
    {
        var app = _app;
        _app = null;
        try { app?.Dispose(); }
        catch (Exception) { /* dead proxy: nothing left to release */ }
    }

    /// <summary>Exposes invalidation for test scenarios that simulate a dead COM proxy.</summary>
    internal void InvalidateForTests() => _executor.Invoke(() => { InvalidateApp(); return 0; });

    /// <summary>
    /// Runs a COM operation on the STA executor with retry and recreate-on-death. Delegates through the IOneNoteApp adapter so the same
    /// path works for both the modern and legacy (v12) COM servers.
    /// </summary>
    private T Dispatch<T>(Func<IOneNoteApp, T> op, int maxAttempts = 5)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        return _executor.Invoke(() => DispatchOnOwner(op, maxAttempts));
    }

    private T DispatchOnOwner<T>(Func<IOneNoteApp, T> op, int maxAttempts)
    {
        // Disposal may have started while this item was queued.
        ObjectDisposedException.ThrowIf(_disposed, this);
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
    }

    /// <summary>Void overload of <see cref="Dispatch{T}"/>.</summary>
    private void Dispatch(Action<IOneNoteApp> op) =>
        Dispatch<int>(a => { op(a); return 0; });

    /// <summary>
    /// Dispatches a NON-idempotent operation with retries disabled (maxAttempts=1).
    /// OneNote's STA COM frequently returns transient RPC_E_CALL_REJECTED /
    /// RPC_E_SERVERCALL_RETRYLATER <em>after</em> a create/new has already taken
    /// effect; retrying such an op would duplicate the created object (e.g. a page).
    /// Serialisation and recreate-on-death behaviour are preserved; only
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

    /// <summary>
    /// Rejects new calls, waits for in-flight work, releases the COM object on its
    /// owner thread, then stops that thread. Idempotent.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposeStarted, 1) == 1) return;
        _disposed = true;
        _executor.Invoke(() => { InvalidateApp(); return 0; });
        _executor.Dispose();
    }
}
