using System.Linq;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;

namespace OneNoteMcp.Tools;

/// <summary>
/// Shared helper that resolves a caller-supplied version token to a CLSID+major pair,
/// capability-gates the request, and routes it to the correct per-CLSID session.
/// Centralised here so every tool gets the same resolution, gating, and error mapping.
/// </summary>
internal static class ToolVersion
{
    /// <summary>
    /// Resolves a version token to (CLSID, Major). Unknown tokens or CLSIDs not
    /// in the catalog throw <see cref="ArgumentException"/>.
    /// </summary>
    public static (string Clsid, int Major) Resolve(string version)
    {
        // VersionResolver.Resolve throws ArgumentException on empty or unknown token.
        var clsid = VersionResolver.Resolve(version);
        var known = OneNoteVersionCatalog.All
            .FirstOrDefault(k => string.Equals(k.Clsid, clsid, StringComparison.OrdinalIgnoreCase));
        if (known is null)
            throw new ArgumentException(
                $"Version '{version}' resolved to CLSID {clsid}, which is not a known OneNote version.",
                nameof(version));
        return (clsid, known.Major);
    }

    /// <summary>
    /// Resolves the version and gates against a required capability BEFORE touching
    /// COM. Throws <see cref="NotSupportedInVersionException"/> when the capability
    /// is absent on that major, so COM is never reached for unsupported calls.
    /// Returns the correctly-keyed session for the CLSID.
    /// </summary>
    public static OneNoteSession Route(string version, Capability cap)
    {
        var (clsid, major) = Resolve(version);
        if (!OneNoteCapabilities.Supports(major, cap))
            throw new NotSupportedInVersionException(major, cap.ToString());
        return OneNoteSession.For(clsid);
    }

    /// <summary>
    /// Route + body wrapped in <see cref="ToolError.Guard"/> so any exception
    /// (ArgumentException, NotSupportedInVersionException, COMException, …) is
    /// mapped to a human-readable error string instead of propagating as an
    /// unhandled exception.
    /// </summary>
    public static string Guarded(string version, Capability cap, Func<OneNoteSession, string> body) =>
        ToolError.Guard(() => body(Route(version, cap)));
}
