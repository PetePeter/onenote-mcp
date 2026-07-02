using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace OneNoteMcp.Interop;

/// <summary>
/// A single catalogued COM/OneNote error: its HRESULT, the OneNote symbolic name,
/// a human-readable message, and a code-specific suggested next action.
/// </summary>
public sealed record ComErrorInfo(int HResult, string Symbol, string Message, string SuggestedAction);

/// <summary>
/// Pure, COM-free translator from raw OneNote/COM HRESULTs to human-readable text
/// plus an actionable hint. Centralised here so every tool surfaces the same
/// wording and callers never see a bare 0x8004xxxx code.
/// </summary>
public static class ComErrorMapper
{
    // Authoritative OneNote HRESULTs (from the OneNote COM SDK) plus the COM/RPC
    // infrastructure codes we can hit when OneNote is busy, closed, or unregistered.
    private static readonly Dictionary<int, ComErrorInfo> _catalog = Build();

    private static Dictionary<int, ComErrorInfo> Build()
    {
        var entries = new[]
        {
            new ComErrorInfo(unchecked((int)0x80042000), "hrMalformedXML",
                "The XML is not well-formed.",
                "Fix the XML you passed to the tool."),
            new ComErrorInfo(unchecked((int)0x80042001), "hrInvalidXML",
                "The XML is invalid.",
                "Correct the OneNote XML and retry."),
            new ComErrorInfo(unchecked((int)0x80042004), "hrSectionDoesNotExist",
                "The section does not exist.",
                "The section ID is stale; refresh IDs via onenote_list_notebooks or onenote_get_hierarchy."),
            new ComErrorInfo(unchecked((int)0x80042005), "hrPageDoesNotExist",
                "The page does not exist.",
                "The page ID is stale; refresh IDs via onenote_get_hierarchy."),
            new ComErrorInfo(unchecked((int)0x80042010), "hrLastModifiedDateDidNotMatch",
                "The page changed since you last read it.",
                "Re-read the page with onenote_get_page, reapply your edit, then update again."),
            new ComErrorInfo(unchecked((int)0x8004200b), "hrSectionReadOnly",
                "The section is read-only.",
                "The section cannot be modified; check its sync state or permissions."),
            new ComErrorInfo(unchecked((int)0x8004200c), "hrPageReadOnly",
                "The page is read-only.",
                "The page cannot be modified; check its sync state or permissions."),
            new ComErrorInfo(unchecked((int)0x80042014), "hrObjectDoesNotExist",
                "The object does not exist.",
                "The object ID was not found (IDs change after a sync); refresh them via onenote_list_notebooks or onenote_get_hierarchy."),
            new ComErrorInfo(unchecked((int)0x80042015), "hrNotebookDoesNotExist",
                "The notebook does not exist.",
                "Open the notebook with onenote_open_notebook, or refresh IDs via onenote_list_notebooks."),
            new ComErrorInfo(unchecked((int)0x80042017), "hrInvalidName",
                "The name is invalid.",
                "Choose a name without illegal filename characters."),
            new ComErrorInfo(unchecked((int)0x8004201a), "hrFileAlreadyExists",
                "A file already exists at the target path.",
                "Choose a new output path or delete the existing file first."),
            new ComErrorInfo(unchecked((int)0x8004201b), "hrSectionEncryptedAndLocked",
                "The section is encrypted and locked.",
                "Unlock the section in OneNote, then retry."),
            new ComErrorInfo(unchecked((int)0x8004201d), "hrNotYetSynchronized",
                "OneNote has not yet synchronized this content.",
                "Wait for OneNote to finish syncing, then retry."),
            new ComErrorInfo(unchecked((int)0x8004201E), "hrLegacySection",
                "The section is in the OneNote 2007 (or earlier) format.",
                "Upgrade it with onenote_convert_section before editing."),
            new ComErrorInfo(unchecked((int)0x8004202D), "hrConvertFailed",
                "The format conversion failed.",
                "OneNote could not convert the section; open it in OneNote and try again."),
            new ComErrorInfo(unchecked((int)0x80042023), "hrTimeOut",
                "The action timed out.",
                "OneNote was busy; retry shortly."),
            new ComErrorInfo(unchecked((int)0x80042030), "hrAppInModalUI",
                "OneNote is showing a dialog and cannot respond to automation.",
                "Dismiss the open OneNote dialog (for example the first-run or sign-in screen), then retry."),

            // COM / RPC infrastructure codes.
            new ComErrorInfo(unchecked((int)0x80010001), "RPC_E_CALL_REJECTED",
                "OneNote is busy and rejected the call.",
                "Retry shortly."),
            new ComErrorInfo(unchecked((int)0x8001010A), "RPC_E_SERVERCALL_RETRYLATER",
                "OneNote asked to retry later (it is busy).",
                "Retry shortly."),
            new ComErrorInfo(unchecked((int)0x800706BA), "RPC_S_SERVER_UNAVAILABLE",
                "The OneNote process is unavailable (closed or restarted).",
                "The session reconnects automatically; retry the call."),
            new ComErrorInfo(unchecked((int)0x80080005), "CO_E_SERVER_EXEC_FAILURE",
                "OneNote failed to start for COM automation.",
                "Open OneNote manually and dismiss any first-run or sign-in screen, then retry."),
        };

        var map = new Dictionary<int, ComErrorInfo>();
        foreach (var e in entries)
            map[e.HResult] = e;
        return map;
    }

    /// <summary>True when the HRESULT is catalogued with a specific message and action.</summary>
    public static bool IsKnown(int hResult) => _catalog.ContainsKey(hResult);

    /// <summary>Returns the catalogued info for an HRESULT, or null when unrecognised.</summary>
    public static ComErrorInfo? Lookup(int hResult) =>
        _catalog.TryGetValue(hResult, out var info) ? info : null;

    /// <summary>
    /// Renders an HRESULT to a one-line, human-readable description. Known codes
    /// include the message and suggested action; unknown codes still carry the hex.
    /// </summary>
    public static string Describe(int hResult)
    {
        var info = Lookup(hResult);
        return info is null
            ? $"OneNote error 0x{hResult:X8}: Unrecognized COM error."
            : $"OneNote error 0x{hResult:X8}: {info.Message} Suggested action: {info.SuggestedAction}";
    }

    /// <summary>
    /// Describes an exception: NotSupportedInVersionException gets a version-aware
    /// message; COM exceptions are mapped by HRESULT; any other exception surfaces
    /// its own message unchanged.
    /// </summary>
    public static string Describe(Exception ex) => ex switch
    {
        NotSupportedInVersionException nsv => DescribeUnsupported(nsv),
        COMException com => Describe(com.HResult),
        _ => ex.Message,
    };

    private static string DescribeUnsupported(NotSupportedInVersionException ex)
    {
        var display = OneNoteVersionCatalog.All
            .FirstOrDefault(k => k.Major == ex.VersionMajor)?.DisplayName
            ?? ex.VersionMajor.ToString();
        return $"Unsupported on OneNote {display}: '{ex.MethodName}' is not available in this version.";
    }
}
