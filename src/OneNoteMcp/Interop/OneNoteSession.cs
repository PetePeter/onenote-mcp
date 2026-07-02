using System.Reflection;
using System.Runtime.ExceptionServices;
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
/// Lazy, recreatable bridge to the OneNote COM Application object.
/// All COM calls are serialised through a single semaphore because the
/// OneNote Application object lives in a single-threaded COM apartment.
/// </summary>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public sealed class OneNoteSession : IDisposable
{
    // HRESULT for "RPC server unavailable" — COM object must be recreated
    private const int HrServerUnavailable = unchecked((int)0x800706BA);

    private static readonly Lazy<OneNoteSession> _instance =
        new(() => new OneNoteSession(CreateRealApp), LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>The process-wide singleton session. Created on first access.</summary>
    public static OneNoteSession Instance => _instance.Value;

    /// <summary>
    /// The most recent COM failure (mapped to human-readable text) seen by any
    /// session in this process, so diagnostics can surface a "last error" without
    /// any logging infrastructure. Null until the first COM failure.
    /// </summary>
    public static string? LastComError { get; private set; }

    private readonly Func<object> _appFactory;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private object? _app;

    /// <summary>
    /// Creates a session with an injectable COM app factory.
    /// Used by tests to verify recreate-on-death behaviour without real COM.
    /// </summary>
    internal OneNoteSession(Func<object> appFactory)
    {
        _appFactory = appFactory;
    }

    private static object CreateRealApp()
    {
        // Early-bound coclass instantiation. The generated interop assembly carries
        // the CLSID, so `new Application()` produces a strongly-typed IApplication
        // RCW. This is the path that actually works against OneNote on .NET Core —
        // the late-bound ProgID/InvokeMember path throws TYPE_E_LIBNOTREGISTERED.
        return new Application();
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
        {
            _app = _appFactory();
        }
    }

    /// <summary>
    /// Marks the current COM app object as dead so <see cref="EnsureApp"/>
    /// will recreate it on the next call.
    /// </summary>
    private void InvalidateApp()
    {
        _app = null;
    }

    /// <summary>
    /// Exposes invalidation for test scenarios that simulate a dead COM proxy.
    /// </summary>
    internal void InvalidateForTests() => InvalidateApp();

    /// <summary>
    /// Runs a COM operation under the serialising semaphore with retry and
    /// recreate-on-death. When the live application is present it is dispatched
    /// strongly-typed via <paramref name="typed"/>; the injected test-fake path
    /// (a duck-typed POCO) is dispatched late-bound via <paramref name="fake"/>.
    /// </summary>
    private T Dispatch<T>(Func<IApplication, T> typed, Func<object, T> fake)
    {
        _lock.Wait();
        try
        {
            return ComRetry.Execute(() =>
            {
                EnsureApp();
                try
                {
                    return _app is IApplication real ? typed(real) : fake(_app!);
                }
                catch (COMException ex)
                {
                    // Record every COM failure for diagnostics before existing handling.
                    LastComError = ComErrorMapper.Describe(ex);
                    if (ex.HResult == HrServerUnavailable)
                    {
                        // Dead proxy — invalidate so EnsureApp recreates next time.
                        InvalidateApp();
                    }
                    throw; // ComRetry sees non-transient and rethrows.
                }
            });
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
    private void Dispatch(Action<IApplication> typed, Action<object> fake) =>
        Dispatch<int>(
            app => { typed(app); return 0; },
            obj => { fake(obj); return 0; });

    /// <summary>
    /// Late-bound invocation used ONLY for injected test fakes. args is passed by
    /// reference so out/ref parameters written back by the fake are visible after
    /// the call. Unwraps TargetInvocationException so the real COMException (with
    /// its HRESULT) reaches ComRetry / the recreate-on-death guard.
    /// </summary>
    private static void InvokeFake(object app, string method, object?[] args)
    {
        try
        {
            app.GetType().InvokeMember(
                method, BindingFlags.InvokeMethod, null, app, args);
        }
        catch (TargetInvocationException tie) when (tie.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            throw; // unreachable; satisfies compiler
        }
    }

    // ── Public COM wrappers ───────────────────────────────────────────────────
    // Public surface is unchanged (int consts in, string out); only the internals
    // switched from late-bound InvokeMember to early-bound IApplication calls.

    /// <summary>
    /// Returns the OneNote hierarchy XML starting from <paramref name="startNodeId"/>.
    /// Pass empty string for the root (all notebooks).
    /// </summary>
    public string GetHierarchy(string startNodeId, int scope, int xmlSchema) =>
        Dispatch(
            real =>
            {
                real.GetHierarchy(startNodeId, (HierarchyScope)scope, out var xml, (XMLSchema)xmlSchema);
                return xml;
            },
            fake =>
            {
                var args = new object?[] { startNodeId, scope, string.Empty, xmlSchema };
                InvokeFake(fake, "GetHierarchy", args);
                return (string)args[2]!;
            });

    /// <summary>Runs a OneNote search and returns hierarchy XML of matching pages.</summary>
    public string FindPages(string searchString, int xmlSchema, string startNodeId = "") =>
        Dispatch(
            real =>
            {
                real.FindPages(startNodeId, searchString, out var xml, false, false, (XMLSchema)xmlSchema);
                return xml;
            },
            fake =>
            {
                var args = new object?[] { startNodeId, searchString, string.Empty, false, false, xmlSchema };
                InvokeFake(fake, "FindPages", args);
                return (string)args[2]!;
            });

    /// <summary>Returns the XML content of a OneNote page.</summary>
    public string GetPageContent(string pageId, int pageInfo, int xmlSchema) =>
        Dispatch(
            real =>
            {
                real.GetPageContent(pageId, out var xml, (PageInfo)pageInfo, (XMLSchema)xmlSchema);
                return xml;
            },
            fake =>
            {
                var args = new object?[] { pageId, string.Empty, pageInfo, xmlSchema };
                InvokeFake(fake, "GetPageContent", args);
                return (string)args[1]!;
            });

    /// <summary>Saves modified page XML back to OneNote.</summary>
    public void UpdatePageContent(string pageChangesXml, int xmlSchema) =>
        Dispatch(
            real => real.UpdatePageContent(pageChangesXml, DateTime.MinValue, (XMLSchema)xmlSchema, true),
            fake => InvokeFake(fake, "UpdatePageContent",
                new object?[] { pageChangesXml, DateTime.MinValue, xmlSchema, true }));

    /// <summary>
    /// Opens or creates a hierarchy node. Returns the object ID of the node.
    /// </summary>
    public string OpenHierarchy(string path, string relativeToObjectId, int createFileType) =>
        Dispatch(
            real =>
            {
                real.OpenHierarchy(path, relativeToObjectId, out var id, (CreateFileType)createFileType);
                return id;
            },
            fake =>
            {
                var args = new object?[] { path, relativeToObjectId, string.Empty, createFileType };
                InvokeFake(fake, "OpenHierarchy", args);
                return (string)args[2]!;
            });

    /// <summary>Publishes a hierarchy node to a file in the given format.</summary>
    public void Publish(string hierarchyId, string targetFilePath, int publishFormat,
        string clsidExporter = "") =>
        Dispatch(
            real => real.Publish(hierarchyId, targetFilePath, (PublishFormat)publishFormat, clsidExporter),
            fake => InvokeFake(fake, "Publish",
                new object?[] { hierarchyId, targetFilePath, publishFormat, clsidExporter }));

    /// <summary>Creates a new page in a section. Returns the new page's object ID.</summary>
    public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault) =>
        Dispatch(
            real =>
            {
                real.CreateNewPage(sectionId, out var id, (NewPageStyle)newPageStyle);
                return id;
            },
            fake =>
            {
                var args = new object?[] { sectionId, string.Empty, newPageStyle };
                InvokeFake(fake, "CreateNewPage", args);
                return (string)args[1]!;
            });

    /// <summary>Closes an open notebook. Does not delete files on disk.</summary>
    public void CloseNotebook(string notebookId) =>
        Dispatch(
            real => real.CloseNotebook(notebookId),
            fake => InvokeFake(fake, "CloseNotebook", new object?[] { notebookId }));

    /// <summary>Deletes a hierarchy node (page/section/notebook) by its object ID. No conflict check.</summary>
    public void DeleteHierarchy(string objectId) =>
        Dispatch(
            real => real.DeleteHierarchy(objectId, DateTime.MinValue, true),
            fake => InvokeFake(fake, "DeleteHierarchy", new object?[] { objectId, DateTime.MinValue, true }));

    /// <summary>Applies a hierarchy-change XML fragment (e.g. a renamed node) to OneNote.</summary>
    public void UpdateHierarchy(string changesXml, int xmlSchema) =>
        Dispatch(
            real => real.UpdateHierarchy(changesXml, (XMLSchema)xmlSchema),
            fake => InvokeFake(fake, "UpdateHierarchy", new object?[] { changesXml, xmlSchema }));

    // ── Full IApplication coverage (P-0543) ───────────────────────────────────
    // QuickFiling is intentionally NOT wrapped: it surfaces the interactive
    // IQuickFilingDialog modal, which has no headless/automation surface.
    // OnNavigate/OnHierarchyChange events are out of scope (unsupported in managed
    // code per MSDN).

    /// <summary>
    /// Fetches a page's binary object (identified by a callback ID from the page
    /// XML) as a base64 string. General binary-by-callback retrieval.
    /// </summary>
    public string GetBinaryPageContent(string pageId, string callbackId) =>
        Dispatch(
            real =>
            {
                real.GetBinaryPageContent(pageId, callbackId, out var b64);
                return b64;
            },
            fake =>
            {
                var args = new object?[] { pageId, callbackId, string.Empty };
                InvokeFake(fake, "GetBinaryPageContent", args);
                return (string)args[2]!;
            });

    /// <summary>Deletes a single content object (by object ID) from a page.</summary>
    public void DeletePageContent(string pageId, string objectId) =>
        Dispatch(
            real => real.DeletePageContent(pageId, objectId, DateTime.MinValue, true),
            fake => InvokeFake(fake, "DeletePageContent",
                new object?[] { pageId, objectId, DateTime.MinValue, true }));

    /// <summary>Returns the object ID of the parent of the given hierarchy node.</summary>
    public string GetHierarchyParent(string objectId) =>
        Dispatch(
            real =>
            {
                real.GetHierarchyParent(objectId, out var parentId);
                return parentId;
            },
            fake =>
            {
                var args = new object?[] { objectId, string.Empty };
                InvokeFake(fake, "GetHierarchyParent", args);
                return (string)args[1]!;
            });

    /// <summary>
    /// Returns the filesystem path of a OneNote special location
    /// (backup folder, unfiled-notes section, or default notebook folder).
    /// </summary>
    public string GetSpecialLocation(int specialLocation) =>
        Dispatch(
            real =>
            {
                real.GetSpecialLocation((SpecialLocation)specialLocation, out var path);
                return path;
            },
            fake =>
            {
                var args = new object?[] { specialLocation, string.Empty };
                InvokeFake(fake, "GetSpecialLocation", args);
                return (string)args[1]!;
            });

    /// <summary>Navigates the OneNote UI to a hierarchy node and optional object.</summary>
    public void NavigateTo(string hierarchyObjectId, string objectId = "", bool newWindow = false) =>
        Dispatch(
            real => real.NavigateTo(hierarchyObjectId, objectId, newWindow),
            fake => InvokeFake(fake, "NavigateTo",
                new object?[] { hierarchyObjectId, objectId, newWindow }));

    /// <summary>Navigates the OneNote UI to a onenote: URL.</summary>
    public void NavigateToUrl(string url, bool newWindow = false) =>
        Dispatch(
            real => real.NavigateToUrl(url, newWindow),
            fake => InvokeFake(fake, "NavigateToUrl", new object?[] { url, newWindow }));

    /// <summary>Returns a onenote: hyperlink to a hierarchy node (and optional page object).</summary>
    public string GetHyperlinkToObject(string hierarchyId, string pageContentObjectId = "") =>
        Dispatch(
            real =>
            {
                real.GetHyperlinkToObject(hierarchyId, pageContentObjectId, out var link);
                return link;
            },
            fake =>
            {
                var args = new object?[] { hierarchyId, pageContentObjectId, string.Empty };
                InvokeFake(fake, "GetHyperlinkToObject", args);
                return (string)args[2]!;
            });

    /// <summary>Returns a web (https) hyperlink to a hierarchy node (and optional page object).</summary>
    public string GetWebHyperlinkToObject(string hierarchyId, string pageContentObjectId = "") =>
        Dispatch(
            real =>
            {
                real.GetWebHyperlinkToObject(hierarchyId, pageContentObjectId, out var link);
                return link;
            },
            fake =>
            {
                var args = new object?[] { hierarchyId, pageContentObjectId, string.Empty };
                InvokeFake(fake, "GetWebHyperlinkToObject", args);
                return (string)args[2]!;
            });

    /// <summary>
    /// Searches page metadata by name and returns hierarchy XML of matching pages.
    /// Pass empty <paramref name="startNodeId"/> to search from the root.
    /// </summary>
    public string FindMeta(string searchName, int xmlSchema, string startNodeId = "") =>
        Dispatch(
            real =>
            {
                real.FindMeta(startNodeId, searchName, out var xml, false, (XMLSchema)xmlSchema);
                return xml;
            },
            fake =>
            {
                var args = new object?[] { startNodeId, searchName, string.Empty, false, xmlSchema };
                InvokeFake(fake, "FindMeta", args);
                return (string)args[2]!;
            });

    /// <summary>Three-way merges OneNote files (base/client/server) into a target file.</summary>
    public void MergeFiles(string baseFile, string clientFile, string serverFile, string targetFile) =>
        Dispatch(
            real => real.MergeFiles(baseFile, clientFile, serverFile, targetFile),
            fake => InvokeFake(fake, "MergeFiles",
                new object?[] { baseFile, clientFile, serverFile, targetFile }));

    /// <summary>Merges the pages of a source section into a destination section.</summary>
    public void MergeSections(string sourceId, string destId) =>
        Dispatch(
            real => real.MergeSections(sourceId, destId),
            fake => InvokeFake(fake, "MergeSections", new object?[] { sourceId, destId }));

    /// <summary>Forces a sync of the given hierarchy node (notebook/section).</summary>
    public void SyncHierarchy(string hierarchyId) =>
        Dispatch(
            real => real.SyncHierarchy(hierarchyId),
            fake => InvokeFake(fake, "SyncHierarchy", new object?[] { hierarchyId }));

    /// <summary>
    /// Sets the section OneNote files a given kind of Outlook item into
    /// (e-mail, contacts, tasks, …).
    /// </summary>
    public void SetFilingLocation(int filingLocation, int filingLocationType, string sectionId) =>
        Dispatch(
            real => real.SetFilingLocation(
                (FilingLocation)filingLocation, (FilingLocationType)filingLocationType, sectionId),
            fake => InvokeFake(fake, "SetFilingLocation",
                new object?[] { filingLocation, filingLocationType, sectionId }));

    /// <inheritdoc/>
    public void Dispose() => _lock.Dispose();
}
