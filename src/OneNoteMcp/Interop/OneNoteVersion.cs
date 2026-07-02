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
}
