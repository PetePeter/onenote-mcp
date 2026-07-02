using Microsoft.Win32;

namespace OneNoteMcp.Interop;

/// <summary>
/// Detected OneNote installation version. Immutable value object.
/// </summary>
public sealed record OneNoteVersion(int Major, string DisplayName)
{
    /// <summary>
    /// Detects the installed OneNote version from the Windows registry.
    /// Tries HKCR\OneNote.Application\CurVer first, then falls back to
    /// checking HKLM\SOFTWARE\Microsoft\Office\{version}\OneNote\InstallRoot.
    /// Returns null when OneNote is not installed or the registry is inaccessible.
    /// </summary>
    public static OneNoteVersion? Detect()
    {
        // Primary: HKCR\OneNote.Application\CurVer default value
        try
        {
            using var key = Registry.ClassesRoot.OpenSubKey(@"OneNote.Application\CurVer");
            var curVer = key?.GetValue(null) as string;
            var parsed = ParseCurVer(curVer);
            if (parsed is not null)
                return parsed;
        }
        catch { /* Registry inaccessible — fall through */ }

        // Fallback: detect by checking known Office version install roots
        int[] knownMajors = [16, 15, 14];
        foreach (int major in knownMajors)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(
                    $@"SOFTWARE\Microsoft\Office\{major}.0\OneNote\InstallRoot");
                if (key is not null)
                    return new OneNoteVersion(major, MapMajorToDisplay(major));
            }
            catch { /* skip */ }
        }

        return null;
    }

    /// <summary>
    /// Parses a CurVer string such as "OneNote.Application.16" into a version.
    /// Returns null for null input or unrecognised format.
    /// Internal to allow direct unit testing without registry access.
    /// </summary>
    internal static OneNoteVersion? ParseCurVer(string? curVer)
    {
        if (string.IsNullOrWhiteSpace(curVer))
            return null;

        // Expected format: "OneNote.Application.<major>"
        var lastDot = curVer.LastIndexOf('.');
        if (lastDot < 0 || lastDot == curVer.Length - 1)
            return null;

        var suffix = curVer[(lastDot + 1)..];
        if (!int.TryParse(suffix, out int major))
            return null;

        return new OneNoteVersion(major, MapMajorToDisplay(major));
    }

    private static string MapMajorToDisplay(int major) => major switch
    {
        12 => "2007",
        14 => "2010",
        15 => "2013",
        16 => "2016/2019/365",
        _  => $"Office {major}",
    };

    /// <summary>
    /// Detects every OneNote desktop install by probing each known CLSID's
    /// LocalServer32 exe path via the injected <paramref name="localServerLookup"/>
    /// (production supplies the real registry reader; tests supply a fake). CLSIDs
    /// that resolve to the same exe are collapsed to a single install, keeping the
    /// most-modern major — this folds the 2010/2013 compat CLSIDs into the modern
    /// desktop entry they actually launch.
    /// </summary>
    public static IReadOnlyList<OneNoteInstall> DetectAll(Func<string, string?> localServerLookup)
    {
        var candidates = new List<OneNoteInstall>();
        foreach (var known in OneNoteVersionCatalog.All)
        {
            var exePath = localServerLookup(known.Clsid);
            if (string.IsNullOrWhiteSpace(exePath))
                continue;

            candidates.Add(new OneNoteInstall(known.Major, known.DisplayName, known.Clsid, exePath));
        }

        // Collapse installs that share an exe path (case-insensitive), keeping the
        // highest major so its own CLSID/DisplayName survive the merge.
        var byExe = new Dictionary<string, OneNoteInstall>(StringComparer.OrdinalIgnoreCase);
        foreach (var candidate in candidates)
        {
            if (!byExe.TryGetValue(candidate.ExePath, out var existing) || candidate.Major > existing.Major)
                byExe[candidate.ExePath] = candidate;
        }

        return byExe.Values.OrderBy(i => i.Major).ToList();
    }

    /// <summary>
    /// Production overload: detects installs using the real Windows registry to read
    /// each CLSID's LocalServer32 exe path. Not unit-tested (IO shim).
    /// </summary>
    public static IReadOnlyList<OneNoteInstall> DetectAll() => DetectAll(ReadLocalServer32);

    /// <summary>
    /// Reads HKLM\SOFTWARE\Classes\CLSID\{clsid}\LocalServer32 (with the WOW6432Node
    /// fallback for 32-bit registrations), stripping command-line arguments and
    /// surrounding quotes to yield the bare exe path. Returns null on any failure.
    /// </summary>
    private static string? ReadLocalServer32(string clsid)
    {
        string[] roots =
        [
            $@"SOFTWARE\Classes\CLSID\{clsid}\LocalServer32",
            $@"SOFTWARE\WOW6432Node\Classes\CLSID\{clsid}\LocalServer32",
        ];

        foreach (var path in roots)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(path);
                if (key?.GetValue(null) is string raw && !string.IsNullOrWhiteSpace(raw))
                    return StripCommandLine(raw);
            }
            catch { /* Registry inaccessible — try next root */ }
        }

        return null;
    }

    /// <summary>
    /// Extracts the bare exe path from a LocalServer32 command line: a quoted path
    /// takes everything between the first pair of quotes; an unquoted value is cut at
    /// the first ".exe" (any following switches are dropped). Returns null for a
    /// malformed value that opens a quote but never closes it.
    /// </summary>
    private static string? StripCommandLine(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.StartsWith('"'))
        {
            var end = trimmed.IndexOf('"', 1);
            return end > 1 ? trimmed[1..end] : null;
        }

        const string exe = ".exe";
        var idx = trimmed.IndexOf(exe, StringComparison.OrdinalIgnoreCase);
        return idx >= 0 ? trimmed[..(idx + exe.Length)] : trimmed;
    }
}

/// <summary>
/// A single detected OneNote desktop install. Immutable value object. Capabilities
/// are queried separately via OneNoteCapabilities.SupportedBy(Major) to keep the
/// Interop layer free of any dependency on Model.
/// </summary>
public sealed record OneNoteInstall(int Major, string DisplayName, string Clsid, string ExePath);

/// <summary>
/// A known desktop OneNote version: its Office major, friendly display name, COM
/// CLSID, and the friendly tokens that resolve to it. <see cref="Clsid"/> exposes
/// the canonical "B"-uppercase-braces string used everywhere as the CLSID identity.
/// </summary>
public sealed record KnownOneNoteVersion(int Major, string DisplayName, Guid ClsidGuid, IReadOnlyList<string> Tokens)
{
    /// <summary>Canonical CLSID string: uppercase, brace-wrapped (e.g. "{0039FFEC-...}").</summary>
    public string Clsid { get; } = ClsidGuid.ToString("B").ToUpperInvariant();
}

/// <summary>
/// Single source of truth for the known desktop OneNote versions, shared by the
/// version resolver and install detection. Lives in Interop so both Interop and
/// Model can use it without introducing a Model→Interop→Model cycle.
/// </summary>
public static class OneNoteVersionCatalog
{
    /// <summary>All known desktop OneNote versions, ordered by major. Wrapped read-only
    /// so the shared single-source catalog cannot be mutated through an array cast.</summary>
    public static readonly IReadOnlyList<KnownOneNoteVersion> All = Array.AsReadOnly(new KnownOneNoteVersion[]
    {
        new(12, "2007", new Guid("0039FFEC-A022-4232-8274-6B34787BFC27"), ["2007", "12"]),
        new(14, "2010", new Guid("D7FAC39E-7FF1-49AA-98CF-A1DDD316337E"), ["2010", "14"]),
        // {DC67E480...} is historically the 2013 desktop CLSID, but on modern
        // machines it is the shared classic-desktop id used by 2016. It is treated
        // as major 16 / "2016"; the "2013"/"15" tokens are aliases for it.
        new(16, "2016", new Guid("DC67E480-C3CB-49F8-8232-60B0C2056C8E"), ["2016", "16", "2013", "15"]),
    });
}
