using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

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

    private readonly Func<object> _appFactory;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private object? _app;
    private Type? _appType;

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
        var type = Type.GetTypeFromProgID("OneNote.Application")
            ?? throw new InvalidOperationException("OneNote.Application ProgID not found.");
        return Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Failed to create OneNote.Application COM instance.");
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
            _appType = _app.GetType();
        }
    }

    /// <summary>
    /// Marks the current COM app object as dead so <see cref="EnsureApp"/>
    /// will recreate it on the next call.
    /// </summary>
    private void InvalidateApp()
    {
        _app = null;
        _appType = null;
    }

    /// <summary>
    /// Exposes invalidation for test scenarios that simulate a dead COM proxy.
    /// </summary>
    internal void InvalidateForTests() => InvalidateApp();

    /// <summary>
    /// Late-bound method invocation on the OneNote Application COM object.
    /// args is passed by reference so out/ref parameters written back by COM
    /// are visible to the caller after the call returns.
    /// </summary>
    private object? Invoke(string method, object?[] args)
    {
        _lock.Wait();
        try
        {
            return ComRetry.Execute(() =>
            {
                EnsureApp();
                try
                {
                    return _appType!.InvokeMember(
                        method,
                        BindingFlags.InvokeMethod,
                        null,
                        _app,
                        args);
                }
                catch (TargetInvocationException tie) when (tie.InnerException is COMException inner)
                {
                    // InvokeMember wraps the real COM exception — unwrap it so
                    // ComRetry and callers see the actual COMException with its HRESULT.
                    ExceptionDispatchInfo.Capture(inner).Throw();
                    throw; // unreachable; satisfies compiler
                }
                catch (COMException ex) when (ex.HResult == HrServerUnavailable)
                {
                    // Dead proxy — invalidate so EnsureApp recreates next time
                    InvalidateApp();
                    throw; // ComRetry will see non-transient and rethrow
                }
            });
        }
        catch (COMException ex) when (ex.HResult == HrServerUnavailable)
        {
            // If ComRetry rethrew a server-unavailable, ensure app is invalidated
            // (inner catch above handles it first, but guard here too)
            InvalidateApp();
            throw;
        }
        finally
        {
            _lock.Release();
        }
    }

    // ── Public COM wrappers ───────────────────────────────────────────────────

    /// <summary>
    /// Returns the OneNote hierarchy XML starting from <paramref name="startNodeId"/>.
    /// Pass empty string for the root (all notebooks).
    /// </summary>
    public string GetHierarchy(string startNodeId, int scope, int xmlSchema)
    {
        // args[2] is the out xml parameter
        var args = new object?[] { startNodeId, scope, string.Empty, xmlSchema };
        Invoke("GetHierarchy", args);
        return (string)args[2]!;
    }

    /// <summary>Runs a OneNote search and returns hierarchy XML of matching pages.</summary>
    public string FindPages(string searchString, int xmlSchema, string startNodeId = "")
    {
        // args[2] is the out hierarchy-xml parameter
        var args = new object?[] { startNodeId, searchString, string.Empty, false, false, xmlSchema };
        Invoke("FindPages", args);
        return (string)args[2]!;
    }

    /// <summary>Returns the XML content of a OneNote page.</summary>
    public string GetPageContent(string pageId, int pageInfo, int xmlSchema)
    {
        var args = new object?[] { pageId, string.Empty, pageInfo, xmlSchema };
        Invoke("GetPageContent", args);
        return (string)args[1]!;
    }

    /// <summary>Saves modified page XML back to OneNote.</summary>
    public void UpdatePageContent(string pageChangesXml, int xmlSchema)
    {
        var args = new object?[] { pageChangesXml, DateTime.MinValue, xmlSchema, true };
        Invoke("UpdatePageContent", args);
    }

    /// <summary>
    /// Opens or creates a hierarchy node. Returns the object ID of the node.
    /// </summary>
    public string OpenHierarchy(string path, string relativeToObjectId, int createFileType)
    {
        var args = new object?[] { path, relativeToObjectId, string.Empty, createFileType };
        Invoke("OpenHierarchy", args);
        return (string)args[2]!;
    }

    /// <summary>Publishes a hierarchy node to a file in the given format.</summary>
    public void Publish(string hierarchyId, string targetFilePath, int publishFormat,
        string clsidExporter = "")
    {
        var args = new object?[] { hierarchyId, targetFilePath, publishFormat, clsidExporter };
        Invoke("Publish", args);
    }

    /// <summary>Creates a new page in a section. Returns the new page's object ID.</summary>
    public string CreateNewPage(string sectionId, int newPageStyle = OneNoteNewPageStyle.NpsDefault)
    {
        var args = new object?[] { sectionId, string.Empty, newPageStyle };
        Invoke("CreateNewPage", args);
        return (string)args[1]!;
    }

    /// <summary>Closes an open notebook. Does not delete files on disk.</summary>
    public void CloseNotebook(string notebookId)
    {
        Invoke("CloseNotebook", new object?[] { notebookId });
    }

    /// <summary>Deletes a hierarchy node (page/section/notebook) by its object ID. No conflict check.</summary>
    public void DeleteHierarchy(string objectId)
    {
        Invoke("DeleteHierarchy", new object?[] { objectId, DateTime.MinValue, true });
    }

    /// <summary>Applies a hierarchy-change XML fragment (e.g. a renamed node) to OneNote.</summary>
    public void UpdateHierarchy(string changesXml, int xmlSchema)
    {
        Invoke("UpdateHierarchy", new object?[] { changesXml, xmlSchema });
    }

    /// <inheritdoc/>
    public void Dispose() => _lock.Dispose();
}
